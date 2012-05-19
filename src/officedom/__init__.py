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
import sys

import win32com.client.gencache
import win32com.client.selecttlb

import word

__all__ = ["word"]

def _init():
    """Initialize office-wide and application type libraries."""
    tlb_modules = {"Office": sys.modules[__name__], "Word": word}
    # Get the type libraries for each supported office application.
    app_entries = tlb_modules.iteritems()
    tlbs = win32com.client.selecttlb.EnumTlbs()
    num_of_tlbs = len(tlbs)
    tlb_suffix = "Object Library"

    for cur_app, cur_mod in app_entries:

        cur_tlb = 0
        tlb_prefix = "Microsoft " + cur_app

        # Generate a module for the current supported office
        # application.
        while cur_tlb < num_of_tlbs and (not tlbs[cur_tlb].desc.startswith(
            tlb_prefix) or not tlbs[cur_tlb].desc.endswith(tlb_suffix)):
            cur_tlb += 1

        if cur_tlb < num_of_tlbs:
            cur_mod.constants = win32com.client.gencache.EnsureModule(
                tlbs[cur_tlb].clsid, tlbs[cur_tlb].lcid, int(tlbs[
                cur_tlb].major), int(tlbs[cur_tlb].minor)).constants

_init()
