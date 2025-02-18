# ***************************************************************************
# *   Copyright (c) 2014 Yorik van Havre <yorik@uncreated.net>              *
# *                                                                         *
# *   This program is free software; you can redistribute it and/or modify  *
# *   it under the terms of the GNU Lesser General Public License (LGPL)    *
# *   as published by the Free Software Foundation; either version 2 of     *
# *   the License, or (at your option) any later version.                   *
# *   for detail see the LICENCE text file.                                 *
# *                                                                         *
# *   This program is distributed in the hope that it will be useful,       *
# *   but WITHOUT ANY WARRANTY; without even the implied warranty of        *
# *   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the         *
# *   GNU Library General Public License for more details.                  *
# *                                                                         *
# *   You should have received a copy of the GNU Library General Public     *
# *   License along with this program; if not, write to the Free Software   *
# *   Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  *
# *   USA                                                                   *
# *                                                                         *
# ***************************************************************************
"""Provide the exporter for SH3D files used above all in Arch and BIM.
"""
## @package exportSH3D
#  \ingroup ARCH
#  \brief SH3D file format exporter
#
#  This module provides tools to export SH3D files.

import FreeCAD

import FreeCADGui
from FreeCAD import Base


__title__  = "FreeCAD SH3D export"
__author__ = ("Julien Masnada")
__url__    = "https://www.freecad.org"

PARAMS = FreeCAD.ParamGet("User parameter:BaseApp/Preferences/Mod/BIM")

DEBUG = True

def export(export_list, filename, colors=None, preferences=None):
    """Export the selected objects to SH3D format.

    Parameters
    ----------
    colors:
        It defaults to `None`.
        It is an optional dictionary of `objName:shapeColorTuple`
        or `objName:diffuseColorList` elements to be used in non-GUI mode
        if you want to be able to export colors.
    """
    import BIM.importers.exportSH3DHelper as exportSH3DHelper
    import BIM.importers.SH3DCommons as SH3DCommons
    if DEBUG:
        from importlib import reload
        reload(exportSH3DHelper)
        reload(SH3DCommons)

    pi = Base.ProgressIndicator()
    try:
        exporter = exportSH3DHelper.SH3DExporter(pi)
        exporter.export_sh3d(export_list, filename, colors)
    finally:
        pi.stop()

    FreeCAD.ActiveDocument.recompute()
