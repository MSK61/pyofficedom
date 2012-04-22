# -*- coding: utf-8 -*-

"""officedom package"""

############################################################
#
# Copyright 2012 Mohammed El-Afifi
# This file is part of pyofficedom.
#
# pyofficedom is free software: you can redistribute it and/or modify
# it under the terms of the GNU Lesser General Public License as
# published by the Free Software Foundation, either version 3 of the
# License, or (at your option) any later version.
#
# pyofficedom is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with pyofficedom.  If not, see
# <http://www.gnu.org/licenses/>.
#
# program:      python office DOM
#
# file:         __init__.py
#
# function:     officedom package
#
# description:  officedom package export file
#
# author:       Mohammed Safwat (MS)
#
# environment:  ActiveState Komodo IDE, version 6.1.3, build 66534,
#               windows xp professional
#
# notes:        This is a private program.
#
############################################################

import itertools
import win32com.client.gencache
import win32com.client.selecttlb

def _init():
    """Initialize application type libraries."""
    app_names = ["Word"]

    # Get the type libraries for supported office applications.
    tlb_prefixes = map(lambda name: "Microsoft " + name, app_names)
    app_tlbs = itertools.ifilter(
        lambda tlb: any(itertools.imap(tlb.desc.startswith, tlb_prefixes)),
        win32com.client.selecttlb.EnumTlbs())

    # Generate modules for supported office applications(if necessary).
    for cur_tlb in app_tlbs:
        win32com.client.gencache.EnsureModule(cur_tlb.clsid, cur_tlb.lcid, int(
            cur_tlb.major), int(cur_tlb.minor))

_init()
