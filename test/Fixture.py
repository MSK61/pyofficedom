# -*- coding: utf-8 -*-

"""contains common test utilities"""

############################################################
#
# Copyright 2012, 2014, 2017 Mohammed El-Afifi
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
# file:         Fixture.py
#
# function:     test utility classes, interfaces, and functions
#
# description:  contains test utilities
#
# author:       Mohammed El-Afifi (ME)
#
# environment:  ActiveState Komodo IDE, version 6.1.3, build 66534,
#               windows xp professional
#
# notes:        This is a private program.
#
############################################################

# an interface-like class
class Fixture:

    """Fixture interface"""

    def setUp(self):
        """Create the context for a test case.

        The method is abstract; it must be overridden in derived
        classes.
        """
        raise NotImplementedError

    def tearDown(self):
        """Clean up the context for a test case.

        The method is abstract; it must be overridden in derived
        classes.
        """
        raise NotImplementedError
