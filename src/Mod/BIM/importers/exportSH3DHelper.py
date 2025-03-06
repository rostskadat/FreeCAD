# ***************************************************************************
# *   Copyright (c) 2024 Julien Masnada <rostskadat@gmail.com>              *
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
"""Helper functions that are used by SH3D exporter."""
import math
import os
import shutil
import uuid
import xml.etree.ElementTree as ET
import xml.dom.minidom as minidom
import zipfile

import BIM.importers.SH3DCommons as SH3DCommons
import Part
from importlib import reload
reload(SH3DCommons)

from draftutils.messages import _wrn, _msg
from draftutils.params import get_param_arch

import FreeCAD as App

if App.GuiUp:
    import FreeCADGui as Gui
    from draftutils.translate import translate
else:
    # \cond
    def translate(_, text):
        return text
    # \endcond

ORIGIN = App.Vector(0, 0, 0)
X_NORM = App.Vector(1, 0, 0)
Y_NORM = App.Vector(0, 1, 0)
Z_NORM = App.Vector(0, 0, 1)
NO_ROT = App.Rotation()

class SH3DExporter:
    """The main class to export an SH3D file.
    """
    def __init__(self, progress_bar=None):
        """Create a SH3DExporter instance to export to the given SH3D file.

        Args:
            progress_bar (ProgressIndicator,optional): a ProgressIndicator
              called to let the User monitor the export process
        """
        super().__init__()
        self.progress_bar = progress_bar
        self.exporters = {
            'Site': self._prepare_site,
            'Building': self._prepare_building,
            'Building Storey': self._prepare_level,
            'Space': self._prepare_room,
            'Wall': self._prepare_wall,
            'Plate': self._prepare_plate,
            'Window': self._prepare_door_or_window,
            'Door': self._prepare_door_or_window,
            'Opening Element': self._prepare_door_or_window,
        }
        self.levels = []
        self.rooms = []
        self.walls = []
        self.walls_by_start = {}
        self.walls_by_end = {}
        self.plates = []
        self.door_or_windows = []
        self.furnitures = []
        self.models = []

    def export_sh3d(self, export_list, filename, colors):
        """Export the active document to the SH3D Home.

        Args:
            home (str): the string containing the XML of the home
                to be imported.

        Raises:
            ValueError: if an invalid SH3D file is detected
        """
        self.export_list = export_list
        self.filename = filename
        self.colors = colors

        if App.GuiUp and get_param_arch("sh3dShowExportDialog"):
            Gui.showPreferences("Import-Export", 7)

        self._get_preferences()

        if self.progress_bar:
            self.progress_bar.start(f"Exporting to SweetHome 3D Home. Please wait ...", -1)
            self.progress_bar.stop()

        _msg(f"Exporting {len(export_list)} objects to '{self.filename}' ...")
        home_attributes = {
            "version":"7400",
            "name": f"{App.ActiveDocument.Label}.sh3d",
            "camera": "topCamera",
            "wallHeight": str(250.0),
            "furnitureSortedProperty": 'NAME'
        }
        self.home = ET.Element("home", attrib=home_attributes)
        for obj in export_list:
            self._prepare_obj_for_export(obj)

        # The order of the following calls is important as it follows the
        # SH3D XML schema...
        self._export_properties()
        self._export_furniture_properties()
        self._export_environment()
        self._export_compass()
        self._export_cameras()
        self._export_levels()
        self._export_rooms()
        self._export_walls()
        self._export_door_or_windows()
        self._export_furnitures()

        tmp_dir = os.path.join(App.ActiveDocument.TransientDir, str(uuid.uuid4()))
        try:
            with zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED) as zip:
                zip.writestr("Home.xml", ET.tostring(self.home, encoding="utf-8"))
                for model in self.models:
                    if model.archive_name in zip.namelist():
                        continue
                    model_path = os.path.join(str(tmp_dir), model.archive_name)
                    os.makedirs(os.path.dirname(model_path), exist_ok=True)
                    _msg(f"Writing model {model.archive_name} to {model_path} ...")
                    model.write(model_path)
                    # zip.write(model_path, model.archive_name)
        finally:
            shutil.rmtree(tmp_dir)

        if True or self.preferences["DEBUG_EXPORT"]:
            xml_str = ET.tostring(self.home, encoding="utf-8")
            pretty_xml = minidom.parseString(xml_str).toprettyxml(indent="  ")
            with open(os.path.join(os.path.dirname(filename), "Home.xml"), "w", encoding="utf-8") as f:
                f.write(pretty_xml)

        _msg(f"Successfully exported to '{self.filename}'")

    def _get_preferences(self):
        """Retrieve the SH3D preferences available in Mod/Arch."""
        self.preferences = {
            'DEBUG_EXPORT': get_param_arch("sh3dDebugExport"),
            'EXPORT_DOORS_AND_WINDOWS': get_param_arch("sh3dExportDoorsAndWindows"),
        }

    def _prepare_obj_for_export(self, obj):
        """Prepare to export an object to the SH3D file.

        Depending upon the type of object, the appropriate method is called to
        prepare the object for export. Note that the IfcType of the object is
        used to determine the appropriate method to call. It is therefore
        important that this attribute be properly set on the object.

        Args:
            obj (any): The object to export.
        """
        if hasattr(obj, "IfcType"):
            _msg(f"Exporting object '{obj.Label}' ({obj.IfcType}) ...")
            self.exporters.get(obj.IfcType, self._not_handled)(obj)
        elif hasattr(obj, "TypeId"):
            if getattr(obj, "TypeId") == "Mesh::Feature":
                _msg(f"Exporting object '{obj.Label}' ({obj.TypeId}) ...")
                self._prepare_furniture(obj)
        if hasattr(obj, "Group"):
            for sub_obj in obj.Group:
                self._prepare_obj_for_export(sub_obj)
        if hasattr(obj, "Proxy") and hasattr(obj.Proxy, "getHosts"):
            for sub_obj in obj.Proxy.getHosts(obj):
                self._prepare_obj_for_export(sub_obj)

    def _prepare_site(self, site):
        self.site = site
        self._export_sh3d_properties(site)

    def _prepare_building(self, building):
        self._export_sh3d_properties(building)

    def _prepare_level(self, level):
        self.levels.append(level)

    def _prepare_room(self, room):
        self.rooms.append(room)

    def _prepare_wall(self, wall):
        """Prepare to export a wall to the SH3D file.

        Store the Arch.Wall to export and its start and end points.

        Args:
            wall (Arch.Wall): The wall to export.
        """
        if wall.Base.TypeId == 'Part::Sweep':
            # This is how the wall are imported...
            start = wall.Base.Spine[0].Start
            end = wall.Base.Spine[0].End
        else:
            start = wall.Base.Start
            end = wall.Base.End
        h_start = SH3DCommons.hash_vector(start)
        h_end = SH3DCommons.hash_vector(end)
        self.walls.append((wall, start, end))
        if h_start not in self.walls_by_start:
            self.walls_by_start[h_start] = []
        if h_end not in self.walls_by_end:
            self.walls_by_end[h_end] = []
        self.walls_by_start[h_start].append(wall.Name)
        self.walls_by_end[h_end].append(wall.Name)

    def _prepare_plate(self, plate):
        self.plates.append(plate)

    def _prepare_door_or_window(self, door_or_window):
        self.door_or_windows.append(door_or_window)

    def _prepare_furniture(self, furniture):
        self.furnitures.append(furniture)
        self.models.append(Model(furniture))

    def _not_handled(self, obj):
        """Generic method to handle unsupported objects.

        Args:
            obj (any): the object to handle.
        """
        _wrn(f"Object '{obj.Label}' with IfcType '{obj.IfcType}' of is not supported. Skipping!")

    def _export_properties(self):
        self._export_property("com.eteks.sweethome3d.SweetHome3D.CatalogPaneDividerLocation", 419)
        self._export_property("com.eteks.sweethome3d.SweetHome3D.ColumnWidths", '151,86,87,91,89')
        self._export_property("com.eteks.sweethome3d.SweetHome3D.FrameHeight", 921)
        self._export_property("com.eteks.sweethome3d.SweetHome3D.FrameWidth", 1381)
        self._export_property("com.eteks.sweethome3d.SweetHome3D.FrameX", 50)
        self._export_property("com.eteks.sweethome3d.SweetHome3D.FrameY", 87)
        self._export_property("com.eteks.sweethome3d.SweetHome3D.PlanPaneDividerLocation", 419)
        self._export_property("com.eteks.sweethome3d.SweetHome3D.PlanViewportX", 0)
        self._export_property("com.eteks.sweethome3d.SweetHome3D.PlanViewportY", 0)
        self._export_property("com.eteks.sweethome3d.SweetHome3D.ScreenHeight", 1152)
        self._export_property("com.eteks.sweethome3d.SweetHome3D.ScreenWidth", 1920)

    def _export_property(self, name, value):
        ET.SubElement(self.home, "property", attrib={"name": name, "value": str(value)})

    def _export_furniture_properties(self):
        self._export_furniture_property('NAME')
        self._export_furniture_property('WIDTH')
        self._export_furniture_property('DEPTH')
        self._export_furniture_property('HEIGHT')
        self._export_furniture_property('VISIBLE')

    def _export_furniture_property(self, name):
        ET.SubElement(self.home, "furnitureVisibleProperty", attrib={"name": name})

    def _export_environment(self):
        attrib = {
            "groundColor": SH3DCommons.color_fc2sh(self.site.groundColor),
            "skyColor": SH3DCommons.color_fc2sh(self.site.skyColor),
            "lightColor": SH3DCommons.color_fc2sh(self.site.lightColor),
            "ceillingLightColor": SH3DCommons.color_fc2sh(self.site.ceillingLightColor),
            "photoWidth": '400',
            "photoHeight": '300',
            "photoAspectRatio": 'VIEW_3D_RATIO',
            "photoQuality": '0',
            "videoWidth": '320',
            "videoAspectRatio": 'RATIO_4_3',
            "videoQuality": '0',
            "videoFrameRate": '25',
        }
        ET.SubElement(self.home, "environment", attrib=attrib)

    def _export_compass(self):
        attrib = {
            "x": '-100.0',
            "y": '100.0',
            "diameter": '100.0',
            "northDirection": '0.0',
            "longitude": '0.0',
            "latitude": '0.0',
            "timeZone": 'Europe/Paris'
        }
        ET.SubElement(self.home, "compass", attrib=attrib)

    def _export_cameras(self):
        attrib = {
            "attribute": 'observerCamera',
            "lens": 'PINHOLE',
            "x": '50.0',
            "y": '50.0',
            "z": '170.0',
            "yaw": '5.4977875',
            "pitch": '0.19634955',
            "fieldOfView": '1.0995575',
            "time": '1739448000000'
        }
        ET.SubElement(self.home, "observerCamera", attrib=attrib)
        attrib = {
            "attribute": 'topCamera',
            "lens": 'PINHOLE',
            "x": '243.99991',
            "y": '994.0298',
            "z": '1125.0',
            "yaw": '3.1415927',
            "pitch": '0.7853982',
            "fieldOfView": '1.0995575',
            "time": '1739448000000'
        }
        ET.SubElement(self.home, "camera", attrib=attrib)

    def _export_levels(self):
        list(map(self._export_level, self.levels))

    def _export_level(self, level):
        attrib = {
            "id": getattr(level, 'id', level.Label),
            "name": level.Label,
            "elevation": str(SH3DCommons.dim_fc2sh(level.Placement.Base.z)),
            "floorThickness": str(SH3DCommons.dim_fc2sh(SH3DCommons.get_fc_property(level, 'floorThickness', 120))),
            "height": str(SH3DCommons.dim_fc2sh(level.Height)),
            "elevationIndex": str(0),
            "visible": "true" #str("true" if level.Visibility else "false")
        }
        ET.SubElement(self.home, "level", attrib)

    def _export_rooms(self):
        list(map(self._export_room, self.rooms))

    def _export_room(self, room):
        attrib = {
            "id": SH3DCommons.get_fc_property(room, 'id', room.Label),
            "level": App.ActiveDocument.getObject(room.ReferenceFloorName).Label,
            "name": room.Label,
        }
        # SH3DCommons.set_sh_attribute(attrib, room, "nameAngle")
        # SH3DCommons.set_sh_attribute(attrib, room, "nameXOffset")
        # SH3DCommons.set_sh_attribute(attrib, room, "nameYOffset")
        SH3DCommons.set_sh_attribute(attrib, room, "areaVisible")
        # SH3DCommons.set_sh_attribute(attrib, room, "areaAngle")
        # SH3DCommons.set_sh_attribute(attrib, room, "areaXOffset")
        # SH3DCommons.set_sh_attribute(attrib, room, "areaYOffset")

        SH3DCommons.set_sh_attribute(attrib, room, "floorVisible", room.ReferenceFloorPanel.Visibility)
        SH3DCommons.set_sh_attribute(attrib, room, "floorColor")
        SH3DCommons.set_sh_attribute(attrib, room, "floorShininess")
        SH3DCommons.set_sh_attribute(attrib, room, "ceilingVisible", room.ReferenceCeilingPanel.Visibility)
        SH3DCommons.set_sh_attribute(attrib, room, "ceilingColor")
        SH3DCommons.set_sh_attribute(attrib, room, "ceilingShininess")
        SH3DCommons.set_sh_attribute(attrib, room, "ceilingShininess")
        SH3DCommons.set_sh_attribute(attrib, room, "ceilingFlat")
        et_room = ET.SubElement(self.home, "room", attrib)
        face = self._get_reference_face(room)
        if face is not None:
            for vertex in face.Vertexes:
                sh_vertex = SH3DCommons.coord_fc2sh(vertex.Point)
                ET.SubElement(et_room, "point", attrib={
                    "x": str(sh_vertex.x),
                    "y": str(sh_vertex.y)
                })

    def _export_walls(self):
        list(map(self._export_wall, self.walls))

    def _export_wall(self, t_wall):
        (wall, start, end) = t_wall
        sh_start = SH3DCommons.coord_fc2sh(start)
        sh_end = SH3DCommons.coord_fc2sh(end)
        attrib = {
            "id": wall.Name,
            "xStart": str(sh_start.x),
            "yStart": str(sh_start.y),
            "xEnd": str(sh_end.x),
            "yEnd": str(sh_end.y),
            "height": str(SH3DCommons.dim_fc2sh(wall.Height)),
            "thickness": str(SH3DCommons.dim_fc2sh(wall.Width)),
            "pattern": 'hatchUp'
        }
        h_start = SH3DCommons.hash_vector(start)
        h_end = SH3DCommons.hash_vector(end)

        # For wallAtStart, priviledges walls that end at that point.
        walls_at_start = self.walls_by_end.get(h_start, [])
        if len(walls_at_start) > 0:
            attrib["wallAtStart"] = walls_at_start[0]
        else:
            walls_at_start = list(filter(lambda w: w != wall.Name, self.walls_by_start.get(h_start, [])))
            if len(walls_at_start) > 0:
                attrib["wallAtStart"] = walls_at_start[0]

        # For wallAtEnd, priviledges walls that start at that point.
        walls_at_end = self.walls_by_start.get(h_end, [])
        if len(walls_at_end) > 0:
            attrib["wallAtEnd"] = walls_at_end[0]
        else:
            walls_at_end = list(filter(lambda w: w != wall.Name, self.walls_by_end.get(h_end, [])))
            if len(walls_at_end) > 0:
                attrib["wallAtEnd"] = walls_at_end[0]

        SH3DCommons.set_sh_attribute(attrib, wall, "pattern")
        SH3DCommons.set_sh_attribute(attrib, wall, "topColor")
        SH3DCommons.set_sh_attribute(attrib, wall, "leftSideColor")
        SH3DCommons.set_sh_attribute(attrib, wall, "leftSideShininess")
        SH3DCommons.set_sh_attribute(attrib, wall, "rightSideColor")
        SH3DCommons.set_sh_attribute(attrib, wall, "rightSideShininess")

        ET.SubElement(self.home, "wall", attrib)

    def _export_door_or_windows(self):
        list(map(self._export_door_or_window, self.door_or_windows))

    def _export_door_or_window(self, door_or_window):
        floor = App.ActiveDocument.getObject(door_or_window.ReferenceFloorName)
        attrib = {
            "id": getattr(door_or_window, 'id', door_or_window.Label),
            "name": door_or_window.Label,
            "level": floor.Label,
        }
        geometry = door_or_window.getPropertyOfGeometry()
        center = geometry.BoundBox.Center
        sh_center = SH3DCommons.coord_fc2sh(center)
        sh_elevation = center-floor.Placement.Base
        sh_elevation.z = sh_elevation.z - geometry.BoundBox.ZLength/2
        sh_elevation = SH3DCommons.coord_fc2sh(sh_elevation)

        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "angle")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "visible")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "movable")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "description")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "information")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "license")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "creator")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "modelMirrored")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "nameVisible")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "nameAngle")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "nameXOffset")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "nameYOffset")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "price")

        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "catalogId")
        SH3DCommons.set_sh_attribute(attrib, door_or_window, "x", sh_center.x)
        SH3DCommons.set_sh_attribute(attrib, door_or_window, "y", sh_center.y)
        SH3DCommons.set_sh_attribute(attrib, door_or_window, "elevation", sh_elevation.z)
        SH3DCommons.set_sh_attribute(attrib, door_or_window, "width") #, SH3DCommons.dim_fc2sh(geometry.BoundBox.XLength))
        SH3DCommons.set_sh_attribute(attrib, door_or_window, "depth") #, SH3DCommons.dim_fc2sh(geometry.BoundBox.YLength))
        SH3DCommons.set_sh_attribute(attrib, door_or_window, "height") #, SH3DCommons.dim_fc2sh(geometry.BoundBox.ZLength))
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "dropOnTopElevation")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "model")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "icon")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "planIcon")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "modelRotation")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "modelCenteredAtOrigin")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "backFaceShown")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "modelFlags")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "modelSize")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "doorOrWindow")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "resizable")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "deformable")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "texturable")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "staircaseCutOutShape")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "color")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "shininess")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "valueAddedTaxPercentage")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "currency")

        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "wallThickness")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "wallDistance")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "wallWidth")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "wallLeft")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "wallHeight")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "wallTop")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "wallCutOutOnBothSides")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "widthDepthDeformable")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "cutOutShape")
        # SH3DCommons.set_sh_attribute(attrib, door_or_window, "boundToWall")

        # ET.SubElement(self.home, "doorOrWindow", attrib)

    def _export_furnitures(self):
        list(map(self._export_furniture, self.furnitures))

    def _export_furniture(self, furniture):
        floor = App.ActiveDocument.getObject(furniture.ReferenceFloorName)
        attrib = {
            "id": getattr(furniture, 'id', furniture.Label),
            "name": furniture.Label,
            "level": floor.Label,
        }
        geometry = furniture.getPropertyOfGeometry()
        center = geometry.BoundBox.Center
        sh_center = SH3DCommons.coord_fc2sh(center)
        sh_elevation = center-floor.Placement.Base
        sh_elevation.z = sh_elevation.z - geometry.BoundBox.ZLength/2
        sh_elevation = SH3DCommons.coord_fc2sh(sh_elevation)

        SH3DCommons.set_sh_attribute(attrib, furniture, "angle")
        SH3DCommons.set_sh_attribute(attrib, furniture, "visible")
        SH3DCommons.set_sh_attribute(attrib, furniture, "movable")
        # SH3DCommons.set_sh_attribute(attrib, furniture, "description")
        # SH3DCommons.set_sh_attribute(attrib, furniture, "information")
        # SH3DCommons.set_sh_attribute(attrib, furniture, "license")
        # SH3DCommons.set_sh_attribute(attrib, furniture, "creator")
        SH3DCommons.set_sh_attribute(attrib, furniture, "modelMirrored")
        # SH3DCommons.set_sh_attribute(attrib, furniture, "nameVisible")
        # SH3DCommons.set_sh_attribute(attrib, furniture, "nameAngle")
        # SH3DCommons.set_sh_attribute(attrib, furniture, "nameXOffset")
        # SH3DCommons.set_sh_attribute(attrib, furniture, "nameYOffset")
        # SH3DCommons.set_sh_attribute(attrib, furniture, "price")
        SH3DCommons.set_sh_attribute(attrib, furniture, "catalogId")
        SH3DCommons.set_sh_attribute(attrib, furniture, "x", sh_center.x)
        SH3DCommons.set_sh_attribute(attrib, furniture, "y", sh_center.y)
        SH3DCommons.set_sh_attribute(attrib, furniture, "elevation", sh_elevation.z)
        SH3DCommons.set_sh_attribute(attrib, furniture, "width", SH3DCommons.dim_fc2sh(geometry.BoundBox.XLength))
        SH3DCommons.set_sh_attribute(attrib, furniture, "depth", SH3DCommons.dim_fc2sh(geometry.BoundBox.YLength))
        SH3DCommons.set_sh_attribute(attrib, furniture, "height", SH3DCommons.dim_fc2sh(geometry.BoundBox.ZLength))
        SH3DCommons.set_sh_attribute(attrib, furniture, "dropOnTopElevation")
        SH3DCommons.set_sh_attribute(attrib, furniture, "model")
        # SH3DCommons.set_sh_attribute(attrib, furniture, "icon")
        SH3DCommons.set_sh_attribute(attrib, furniture, "modelRotation")
        SH3DCommons.set_sh_attribute(attrib, furniture, "modelCenteredAtOrigin")
        # SH3DCommons.set_sh_attribute(attrib, furniture, "backFaceShown")
        # SH3DCommons.set_sh_attribute(attrib, furniture, "modelFlags")
        # SH3DCommons.set_sh_attribute(attrib, furniture, "modelSize")
        SH3DCommons.set_sh_attribute(attrib, furniture, "doorOrWindow")
        SH3DCommons.set_sh_attribute(attrib, furniture, "resizable")
        SH3DCommons.set_sh_attribute(attrib, furniture, "deformable")
        SH3DCommons.set_sh_attribute(attrib, furniture, "texturable")
        # SH3DCommons.set_sh_attribute(attrib, furniture, "staircaseCutOutShape")
        SH3DCommons.set_sh_attribute(attrib, furniture, "color")
        SH3DCommons.set_sh_attribute(attrib, furniture, "shininess")
        # SH3DCommons.set_sh_attribute(attrib, furniture, "valueAddedTaxPercentage")
        # SH3DCommons.set_sh_attribute(attrib, furniture, "currency")
        SH3DCommons.set_sh_attribute(attrib, furniture, "horizontallyRotatable")
        SH3DCommons.set_sh_attribute(attrib, furniture, "pitch")
        SH3DCommons.set_sh_attribute(attrib, furniture, "roll")
        # SH3DCommons.set_sh_attribute(attrib, furniture, "widthInPlan")
        # SH3DCommons.set_sh_attribute(attrib, furniture, "depthInPlan")
        # SH3DCommons.set_sh_attribute(attrib, furniture, "heightInPlan")
        ET.SubElement(self.home, "pieceOfFurniture", attrib)

    def _get_reference_face(self, obj):
        if hasattr(obj, "ReferenceBottomFace"):
            return obj.ReferenceBottomFace.Shape
        for face in obj.Shape.Faces:
            if face.normalAt(0, 0).getAngle(Z_NORM) < 0.01:
                return face
        return None

    def _export_sh3d_properties(self, obj):
        """Generic method to export properties of an object to SH3D.

        Args:
            obj (FreeCAD Object): the object for which to export the given property
        """
        for property in obj.PropertiesList:
            if obj.getGroupOfProperty(property) == "SweetHome3D":
                _msg(f"Exporting property '{property}' of element '{obj.Name}' ...")

class Model:
    # This encapsulate the export of a Furniture / Mesh

    def __init__(self, furniture):
        self.furniture = furniture

    def __getattribute__(self, name: str, /):
        if name == "archive_name":
            return SH3DCommons.get_fc_property(self.furniture, "model")
        return super().__getattribute__(name)

    def write(self, path):
        mesh = self.furniture.getPropertyOfGeometry().copy()
        # Apply the required transformations in reverse order
        # Look at importSH3DHelper.FurnitureHandler._create_furniture
        # for details...
        mesh_transform = App.Matrix()
        if self.furniture.angle != 0:
            mesh_transform = App.Rotation(Z_NORM, Degree=-self.furniture.angle).toMatrix().multiply(mesh_transform)
        if self.furniture.roll != 0:
            mesh_transform = App.Rotation(Y_NORM, Degree=-self.furniture.roll).toMatrix().multiply(mesh_transform)
        if self.furniture.pitch != 0:
            mesh_transform = App.Rotation(X_NORM, Degree=self.furniture.pitch).toMatrix().multiply(mesh_transform)

        model_transform = self._apply_scale_invariant(App.Matrix())

        model_bb = mesh.BoundBox
        normalized_model = Part.makeBox(model_bb.XLength, model_bb.YLength, model_bb.ZLength)
        normalized_model = normalized_model.transformGeometry(model_transform)
        normilized_bb = normalized_model.BoundBox
        # Note it is the reverse from the import transformation.
        x_scale = self.furniture.getPropertyOfGeometry().BoundBox.XLength / normilized_bb.XLength
        y_scale = self.furniture.getPropertyOfGeometry().BoundBox.YLength / normilized_bb.YLength
        z_scale = self.furniture.getPropertyOfGeometry().BoundBox.ZLength / normilized_bb.ZLength

        mesh_transform.scale(x_scale, y_scale, z_scale)
        # mesh_transform = self._apply_scale_invariant(mesh_transform)
        # mesh.transform(mesh_transform)
        mesh.write(path, 'OBJ')

    def _apply_scale_invariant(self, transform):
        transform.rotateX(-math.pi/2)
        if self.furniture.modelMirrored:
            transform.scale(-1, 1, 1) # Mirror along X
        if self.furniture.modelRotation:
            rij = [ float(v) for v in self.furniture.modelRotation.split() ]
            rotation = App.Matrix(
                App.Vector(rij[0], rij[3], rij[6]),
                App.Vector(rij[1], rij[4], rij[7]),
                App.Vector(rij[2], rij[5], rij[8])
            )
            transform = rotation.inverse().multiply(transform)
        transform.move(-self.furniture.getPropertyOfGeometry().BoundBox.Center)
        return transform
