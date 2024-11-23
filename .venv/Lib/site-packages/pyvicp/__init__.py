# -*- coding: utf-8 -*-
"""Pure Python implementation of LeCroy VICP protocol


:copyright: 2003-2022 by Anthony Cake, Urs Schroffenegger, and Bob McNamara,
:see AUTHORS for more details.
:
:license: LGPL-2.1-or-later, see LICENSE for more details.
"""
from .error_status_register import EXR_LOOKUP
from .version import __version__
from .vicpclient import SERVER_PORT_NUM, Client, ProtocolError
