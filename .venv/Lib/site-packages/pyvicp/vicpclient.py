"""
    Summary:    Lightweight VICP client implementation.

    Started by: Anthony Cake, June 2003

    Python rewrite by: Bob McNamara, December 2016

        Published on SourceForge under LeCroyVICP project, Sept 2003

        This library is free software; you can redistribute it and/or
        modify it under the terms of the GNU Lesser General Public
        License as published by the Free Software Foundation; either
        version 2.1 of the License, or (at your option) any later version.

        This library is distributed in the hope that it will be useful,
        but WITHOUT ANY WARRANTY; without even the implied warranty of
        MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
        Lesser General Public License for more details.

        You should have received a copy of the GNU Lesser General Public
        License along with this library; if not, write to the Free Software
        Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

    ------------------------------------------------------------------------------

    Description:

        This file contains a Client-side implementation of the VICP
        network communications protocol used to control LeCroy Digital
        Oscilloscopes (DSOs).

    VICP Protocol Description/History:

        The VICP Protocol has been around since 1997/98. It did not
        change in any way between its conception and June 2003, when
        a previously reserved field in the header was assigned.  This
        field, found at byte #2, is now used to allow the client-end
        of a VICP communication to detect 'out of sync' situations,
        and therefore allows the IEEE 488.2 'Unread Response' mechanism
        to be emulated.

        These extensions to the original protocol did not cause a
        version number change, and are referred to as version 1a. It
        was decided not to bump the version number to reduce the impact
        on clients, many of which are looking for a version number of
        1.  Clients and servers detect protocol version 1a by examining
        the sequence number field, it should be 0 for version 1 of the
        protocol (early clients), or non-zero for v1a.

    VICP Header is always exactly 8 bytes long:

        Byte   Description
        -------------------------------------------
        0      Operation
        1      Version         1 = version 1
        2      Sequence Number { 1..255 }, (was unused until June 2003)
        3      Unused
        4      Block size, MSB  (not including this header)
        5      Block size
        6      Block size
        7      Block size, LSB

        Byte 0 (Operation) bit definitions:

        Bit    Mnemonic      Purpose
        ------------------------------------
        D7     DATA          Data block (D0 indicates with/without EOI)
        D6     REMOTE        Remote Mode
        D5     LOCKOUT       Local Lockout (Lockout front panel)
        D4     CLEAR         Device Clear (if sent with data, clear occurs
                                           before block is passed to parser)
        D3     SRQ           SRQ (Device -> PC only)
        D2     SERIALPOLL    Request a serial poll
        D1     Reserved      Reserved for future expansion
        D0     EOI           Block terminated in EOI

    Known Limitations:

    Outstanding Issues

    Dependencies
        - Uses generic Python sockets, so should run anywhere Python runs.
"""

import logging
import socket
import struct
import time
from dataclasses import dataclass
from typing import Optional

logger = logging.getLogger(__name__)

SERVER_PORT_NUM = 1861  # port # registered with IANA for lecroy-vicp
HEADER_FORMAT = "!BBBBI"  # format of network header
HEADER_VERSION1 = 0x01  # Header Version
CONNECT_TIMEOUT_SECS = 5
DEFAULT_TIMEOUT_SECS = 10


class OPERATION:
    """VICP header 'Operation' bits"""

    data = 0x80
    remote = 0x40
    lockout = 0x20
    clear = 0x10
    srq = 0x08
    reqserialpoll = 0x04
    undefined = 0x02  # presumably this is reserved for future use
    eoi = 0x01


# ------------------------------------------------------------------------------
# Exceptions


class ProtocolError(Exception):
    """raised when a protocol error is detected"""


class SrqStateChanged(Exception):
    """raised when a change in SRQ state is detected"""


# ------------------------------------------------------------------------------


@dataclass
class RecvBlock:
    """variables associated with receiving blocks"""

    read_state: int  # current state of read 'state machine'
    flags: int  # flags byte from received header
    block_size: int  # total number of payload bytes
    bytes_read: int  # number of payload bytes received so far
    eoi_terminated: bool  # flag indicating whether current block
    srq_state: int  # indicates state of SRQ
    srq_state_changed: bool  # indicates a change in SRQ state
    seq_num: int  # received sequence number


class READSTATE:
    """enum class defining states of read FSM"""

    header = 0
    data = 1


class Client:
    """connection to client device via VICP protocol"""

    def __init__(
        self,
        address: str,
        port: int = SERVER_PORT_NUM,
        timeout: Optional[float] = None,
        debug: bool = False,
    ) -> None:
        if debug:
            logger.setLevel("DEBUG")

        connect_timeout = timeout or CONNECT_TIMEOUT_SECS
        self.ipaddr = address  # IP address of the instrument
        self.port = port  # port number
        self._remote_mode = False  # if True, device is in remote mode
        self._local_lockout = False  # if True, device is in local lockout mode
        self._next_sequence_number = 1  # next sequence value
        self._last_sequence_number = 0  # last used sequence value
        self._flush_unread_responses = True  # if True, unread responses are flushed
        #  (emulate IEEE 488.2 behaviour)
        self._vicpversion1a_supported = (
            False  # version 1a of the VICP protocol supported
        )
        # (seq. numbers and OOB data)
        self._recv_block = RecvBlock(
            read_state=READSTATE.header,  # current state of read 'state machine'
            flags=0,  # flags byte from received header
            block_size=0,  # total number of payload bytes
            bytes_read=0,  # number of payload bytes received so far
            eoi_terminated=False,  # flag indicating whether current block
            srq_state=False,  # indicates state of SRQ
            srq_state_changed=False,  # indicates a change in SRQ state
            seq_num=0,  # received sequence number
        )

        # finally, connect and configure the default timeout
        self._socket = self._connect_to_device(connect_timeout)

        # initialize variables
        self.keepalive = False
        self.timeout = DEFAULT_TIMEOUT_SECS

    @property
    def timeout(self) -> Optional[float]:
        return self._socket.gettimeout()

    @timeout.setter
    def timeout(self, value: float) -> None:
        self._socket.settimeout(value)

    @property
    def keepalive(self) -> bool:
        """Status of the TCP keepalive.

        Keepalive is on/off for both the sync and async sockets

        If a connection is dropped as a result of “keepalives”, the error code
        VI_ERROR_CONN_LOST is returned to current and subsequent I/O
        calls on the session.

        """
        return self._keepalive

    @keepalive.setter
    def keepalive(self, keepalive: bool) -> None:
        self._keepalive = bool(keepalive)
        self._socket.setsockopt(socket.SOL_SOCKET, socket.SO_KEEPALIVE, bool(keepalive))

    # --------------------------------------------------------------------------
    def close(self) -> None:
        """
        close the connection
        """
        logger.debug("Disconnecting:")

        # check if connected
        if hasattr(self, "_socket"):
            logger.debug("close %s:", self._socket)
            try:
                self._socket.shutdown(socket.SHUT_RDWR)
                self._socket.close()
            except OSError:
                pass
            del self._socket

        # reset any partial read operation
        self._recv_block.read_state = READSTATE.header
        self._vicpversion1a_supported = False

    # --------------------------------------------------------------------------
    def _connect_to_device(
        self, timeout: float = CONNECT_TIMEOUT_SECS
    ) -> socket.socket:
        """
        connect to a network device

        address is extracted from self.ipaddr (specified during construction of
        base class)
        """
        # if already connected to scope...
        if hasattr(self, "_socket"):
            return self._socket

        # create client's socket
        logger.debug("Opening Socket:")
        sock = socket.create_connection((self.ipaddr, self.port), timeout=timeout)
        logger.debug("Socket opened: '%s'", sock)

        # disable the TCP/IP 'Nagle' algorithm that buffers packets before sending.
        sock.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)

        # enable SO_LINGER to allow hard closure of sockets
        l_onoff = 1  # LINGER enabled
        l_linger = 0  # timeout = 0
        enable_linger = struct.pack("ii", l_onoff, l_linger)
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_LINGER, enable_linger)

        return sock

    # --------------------------------------------------------------------------
    def device_clear(self, force_reconnect: bool = False) -> None:
        """clear the device"""
        try:
            self._send_packet(flags=OPERATION.clear)
        except socket.error:
            force_reconnect = True  # failed to send, so try to reconnect

        # if VICP version 1a is not supported (which means that
        # unread response clearing is not supported), then momentarily
        # disconnect from the device in order to clear buffers also
        # do this if a reconnection was forced
        if not self._vicpversion1a_supported or force_reconnect:
            # TODO: remove when 'RebootScope' bug is fixed
            time.sleep(100)
            self.close()
            self._socket = self._connect_to_device()

    # --------------------------------------------------------------------------
    def serial_poll(self) -> int:
        """
        return the serial poll byte.  Use the new Out-Of-Band signaling
        technique if supported, else the original 'in-band' technique.
        """
        if self._vicpversion1a_supported:
            return self._oob_data_request(ord("S"))  # 'S' == Serial Poll

        # request the serial poll using an in-band technique
        self._send_packet(flags=OPERATION.reqserialpoll)

        # read the serial-poll response
        serial_poll_response = self.recv_block()
        return serial_poll_response[0]

    # --------------------------------------------------------------------------
    def _oob_data_request(self, request_type: int) -> int:
        """
        out-of band data request
        presently used only for serial polling.
        """
        oob_data_test = struct.pack("!b", request_type)
        self._socket.sendall(oob_data_test, socket.MSG_OOB)
        return self._socket.recv(1, socket.MSG_OOB)[0]

    # --------------------------------------------------------------------------
    def _send_packet(self, payload: bytes = b"", flags: int = OPERATION.eoi) -> None:
        """
        send a block of data to the device.
        """

        # build the header
        eoi_termination = (flags & OPERATION.eoi) != 0
        seq_num = self._get_next_sequence_number(eoi_termination)  # sequence number
        logger.debug("_send_packet: seq=%d eoi=%d", seq_num, eoi_termination)

        flags |= OPERATION.data
        payload_length = len(payload)
        msg = bytearray(
            struct.pack(
                HEADER_FORMAT, flags, HEADER_VERSION1, seq_num, 0, payload_length
            )
        )
        msg.extend(payload)
        self._socket.sendall(msg)

        logger.debug(
            "_send_packet: bytes_sent=%d [%.20s]",
            payload_length,
            payload.decode("utf-8", errors="ignore"),
        )

    # --------------------------------------------------------------------------
    def _get_next_sequence_number(self, eoi_termination: bool) -> int:
        """
        Return the next sequence number in the range 1..255
        (Note that zero is omitted intentionally)

        used to synchronize write/read operations, attempting to
        emulate the 488.2 'discard unread response' behaviour
        """

        # we'll return the current sequence number
        self._last_sequence_number = self._next_sequence_number

        # which then gets incremented if this block is EOI terminated
        if eoi_termination:
            self._next_sequence_number += 1
            if self._next_sequence_number >= 256:
                self._next_sequence_number = 1

        return self._last_sequence_number

    # --------------------------------------------------------------------------
    def send(self, cmdstr: bytes) -> None:
        self._send_packet(cmdstr)

    # --------------------------------------------------------------------------
    def receive(self, buf_len: int = None) -> bytes:
        return self._read_till_whenever(buf_len, stop_at_eob=False)

    # --------------------------------------------------------------------------
    def recv_block(self, buf_len: int = None) -> bytes:
        return self._read_till_whenever(buf_len, stop_at_eob=True)

    # --------------------------------------------------------------------------
    def _receive_flush(self, bytes_to_dump: int) -> None:
        """
        dump data until the next header is found.
        flush exactly 'bytes_to_dump' bytes from socket.
        received data is thrown away and nothing is returned
        """
        logger.debug("_receive_flush: unread response, dumping %d bytes", bytes_to_dump)

        self._receive_exact(bytes_to_dump)

    # --------------------------------------------------------------------------
    def _receive_exact(self, recv_len: int) -> bytes:
        """
        receive exactly 'recv_len' bytes from socket.
        returns a bytearray containing the received data.
        """
        recv_buffer = bytearray(recv_len)
        self._receive_exact_into(memoryview(recv_buffer))
        return recv_buffer

    # --------------------------------------------------------------------------
    def _receive_exact_into(self, view: memoryview) -> None:
        # bytearray) -> None:
        """
        receive data from socket to exactly fill buffer.
        """
        recv_len = len(view)
        bytes_recvd = 0

        while bytes_recvd < recv_len:
            request_size = recv_len - bytes_recvd
            data_len = self._socket.recv_into(view, request_size)
            bytes_recvd += data_len
            view = view[data_len:]

        if bytes_recvd > recv_len:
            logger.error("socket.recv_into scribbled past end of buffer")
            raise MemoryError("socket.recv_into scribbled past end of buffer")

    # --------------------------------------------------------------------------
    def _read_till_whenever(self, buf_len: Optional[int], stop_at_eob: bool) -> bytes:
        # use default length if None given
        req_len = buf_len or 4096

        buf_list = list()

        while True:
            # loop until one of the following occurs:
            #     1. we have read the requested number of bytes (buf_len)
            #     2. stop_at_eob is True and we're at the end of a block
            #     3. _recv_block.eoi_terminated is True
            recv_buffer = self._read_from_device(req_len, stop_at_eob=stop_at_eob)
            nbytes = len(recv_buffer)
            if nbytes < req_len:
                # truncate the buffer to the actual number of bytes received
                recv_buffer = recv_buffer[:nbytes]
            buf_list.append(recv_buffer)

            if buf_len is not None:
                # we've either read the requested number of bytes or hit a stop
                # condition.  in either case, we're done.  exit the loop.
                break

            if self._recv_block.read_state == READSTATE.header:
                # we're at the end of a data block...
                if stop_at_eob or self._recv_block.eoi_terminated:
                    break

        if len(buf_list) > 1:
            # multiple buffers, they need to be concatenated
            result = b"".join(buf_list)
        else:
            # only one buffer; it's our result
            result = buf_list[0]

        return result

    # --------------------------------------------------------------------------
    def _read_from_device(
        self, req_len: int, stop_at_eob: bool = False, return_after_srq: bool = False
    ) -> bytes:
        """
        read block of data from a network device
        """
        reply_buf = bytearray(req_len)
        user_buffer_size_bytes = len(reply_buf)
        user_buffer_bytes_written = 0  # number of bytes placed in user buffer
        view = memoryview(reply_buf)

        while True:
            # loop until one of the following occurs:
            #     1. srq_state_changed is True and return_after_srq is True
            #     2. stop_at_eob is True and we're at the end of a block
            #     3. _recv_block.eoi_terminated is True and we're at the end of a block
            #     4. we've read as many bytes as requested
            if self._recv_block.read_state == READSTATE.header:
                try:
                    self._read_next_header()
                except SrqStateChanged:
                    # if we saw SRQ come (or go), and the user requests that we
                    # return immediately after seeing the change, then return
                    if return_after_srq:
                        break

            if self._recv_block.read_state == READSTATE.data:
                # fill the user-supplied buffer (but no more than
                # src_bytes_available bytes)
                user_buffer_free_space = (
                    user_buffer_size_bytes - user_buffer_bytes_written
                )
                src_bytes_available = (
                    self._recv_block.block_size - self._recv_block.bytes_read
                )

                recv_len = min(user_buffer_free_space, src_bytes_available)

                self._receive_exact_into(view[:recv_len])

                self._recv_block.bytes_read += recv_len
                user_buffer_bytes_written += recv_len
                view = view[recv_len:]

                if self._recv_block.bytes_read >= self._recv_block.block_size:
                    # we have finished reading the contents of this
                    # header-prefixed block.  go back to the state where we can
                    # watch for the next block.
                    self._recv_block.read_state = READSTATE.header

                    if self._recv_block.eoi_terminated:
                        break

                    if stop_at_eob:
                        # stop at end of block
                        break

                if user_buffer_bytes_written >= user_buffer_size_bytes:
                    # the user's buffer has been filled.  get out of the loop.
                    break

        if user_buffer_bytes_written < req_len:
            reply_buf = reply_buf[:user_buffer_bytes_written]

        logger.debug("_read_from_device: returning %d bytes", user_buffer_bytes_written)

        return reply_buf

    # --------------------------------------------------------------------------
    def _read_srq_packet(self) -> None:
        """
        read the SRQ packet, process it, and toss it.
        errors will raise an exception.
        """
        assert self._recv_block.read_state == READSTATE.data
        srq_packet = self._receive_exact(self._recv_block.block_size)
        self._recv_block.read_state = READSTATE.header
        # '1' = asserted, '0' = deasserted
        self._recv_block.srq_state = srq_packet[0] == "1"

    # --------------------------------------------------------------------------
    def _read_next_header(self) -> None:
        """
        read the next header, flushing any unread (old) responses along the way.
        errors will raise an exception.
        """
        while self._recv_block.read_state == READSTATE.header:
            self._recv_block = self._get_recv_block()

            # header was successfully read
            logger.debug(
                "Read Header: "
                "block_size=%d, "
                "EOI=%d, "
                "SRQ state changed=%d, "
                "seq_num=%d",
                self._recv_block.block_size,
                self._recv_block.eoi_terminated,
                self._recv_block.srq_state_changed,
                self._recv_block.seq_num,
            )

            if self._recv_block.srq_state_changed:
                self._read_srq_packet()
                raise SrqStateChanged

            if self._recv_block.seq_num == 0:
                # we're talking to a scope running pre-June 2003 code that
                # didn't support sequence numbering, and therefore we don't
                # know when to dump data.
                self._vicpversion1a_supported = False
            else:
                # version '1a' of the VICP protocol is in use, which in
                # addition to sequence numbering supports the use of
                # out-of-band signaling.
                self._vicpversion1a_supported = True

                # if we're flushing unread responses, and this header contains
                # an unexpected sequence number (older than the current one),
                # then dump this block and go around again.
                raw_delta = self._last_sequence_number - self._recv_block.seq_num
                seq_num_delta = raw_delta + 255 if raw_delta < -128 else raw_delta
                if self._flush_unread_responses and seq_num_delta > 0:
                    self._receive_flush(self._recv_block.block_size)
                    self._recv_block.read_state = READSTATE.header

    # --------------------------------------------------------------------------
    def _get_recv_block(self) -> RecvBlock:
        """
        read one header
        assumes that READSTATE is 'header'.
        errors will raise an exception
        """
        # receive the scope's response, header first
        header_size = struct.calcsize(HEADER_FORMAT)
        try:
            header_buf = self._receive_exact(header_size)
        except socket.timeout:
            # a timeout is usually due to a bad query
            # ...no need to disconnect
            raise
        except socket.error:
            # something more serious than a timeout. since we are
            # out of sync, need to close & reopen the socket
            self.close()
            self._socket = self._connect_to_device()
            raise

        flags, version, seq_num, _, block_size = struct.unpack(
            HEADER_FORMAT, header_buf
        )

        # check the integrity of the header
        if not ((flags & OPERATION.data) and (version == HEADER_VERSION1)):
            error_message_format = (
                "Invalid Header! (header received" + " %02x" * header_size + ")"
            )
            error_message = error_message_format % tuple(header_buf)
            logger.error(error_message)

            # error state, cannot recognise header. since we are
            # out of sync, need to close & reopen the socket
            self.close()
            self._socket = self._connect_to_device()
            raise ProtocolError(error_message)

        return RecvBlock(
            read_state=READSTATE.data,
            flags=flags,
            block_size=block_size,
            bytes_read=0,
            eoi_terminated=flags & OPERATION.eoi,
            srq_state=self._recv_block.srq_state,
            srq_state_changed=flags & OPERATION.srq,
            seq_num=seq_num,
        )
