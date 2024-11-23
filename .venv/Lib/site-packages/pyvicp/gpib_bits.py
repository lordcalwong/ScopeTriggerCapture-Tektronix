# GPIB status bit vector :
# ibsta and wait mask
ERR = 1 << 15  # Error detected                    0x8000
TIMO = 1 << 14  # Timeout                          0x4000
END = 1 << 13  # EOI or EOS detected               0x2000
SRQI = 1 << 12  # SRQ detected by CIC              0x1000
RQS = 1 << 11  # Device needs service              0x0800
SPOLL = 1 << 10  # Board has been serially polled  0x0400
CMPL = 1 << 8  # I/O completed                     0x0100
REM = 1 << 6  # Remote state                       0x0040
CIC = 1 << 5  # Controller-in-Charge               0x0020
ATN = 1 << 4  # Attention asserted                 0x0010
TACS = 1 << 3  # Talker active                     0x0008
LACS = 1 << 2  # Listener active                   0x0004
DTAS = 1 << 1  # Device trigger state              0x0002
DCAS = 1 << 0  # Device clear state                0x0001

# GPIB error codes :
# iberr
EDVR = 0  # System error
ECIC = 1  # Function requires GPIB board to be CIC
ENOL = 2  # Write function detected no Listeners
EADR = 3  # Interface board not addressed correctly
EARG = 4  # Invalid argument to function call
ESAC = 5  # Function requires GPIB board to be SAC
EABO = 6  # I/O operation aborted
ENEB = 7  # Non-existent interface board
EDMA = 8  # Error performing DMA
EOIP = 10  # I/O operation started before previous operation completed
ECAP = 11  # No capability for intended operation
EFSO = 12  # File system operation error
EBUS = 14  # Command error during device call
ESTB = 15  # Serial poll status byte lost
ESRQ = 16  # SRQ remains asserted
ETAB = 20  # The return buffer is full.
ELCK = 21  # Address or board is locked.
