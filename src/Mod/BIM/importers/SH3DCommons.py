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
"""Helper functions that are used by SH3D importer/exporter."""
import FreeCAD as App
import math
import xml.etree.ElementTree as ET

from draftutils.messages import _wrn, _msg, _log

# SweetHome3D is in cm while FreeCAD is in mm
FACTOR = 10
TOLERANCE = float(.1)

def ang_sh2fc(angle:float):
    """Convert SweetHome angle (º) to FreeCAD angle (º)

    SweetHome angles are clockwise positive while FreeCAD are anti-clockwise
    positive

    Args:
        angle (float): The SweetHome angle

    Returns:
        float: the FreeCAD angle
    """
    return -float(angle)


def coord_fc2sh(vector):
    """Converts FreeCAD coordinate to SweetHome coordinate.

    Args:
        FreeCAD.Vector (FreeCAD.Vector): The FreeCAD coordinate

    Returns:
        FreeCAD.Vector: the SweetHome coordinate
    """
    return App.Vector(vector.x/FACTOR, -vector.y/FACTOR, vector.z/FACTOR)


def coord_sh2fc(vector):
    """Converts SweetHome coordinate to FreeCAD coordinate

    Args:
        FreeCAD.Vector (FreeCAD.Vector): The SweetHome coordinate

    Returns:
        FreeCAD.Vector: the FreeCAD coordinate
    """
    return App.Vector(vector.x*FACTOR, -vector.y*FACTOR, vector.z*FACTOR)


def dim_fc2sh(dimension):
    """Convert FreeCAD dimension (mm) to SweetHome dimension (cm)

    Args:
        dimension (float): The FreeCAD dimension

    Returns:
        float: the SweetHome dimension
    """
    return float(dimension)/FACTOR


def dim_sh2fc(dimension):
    """Convert SweetHome dimension (cm) to FreeCAD dimension (mm)

    Args:
        dimension (float): The SweetHome dimension

    Returns:
        float: the FreeCAD dimension
    """
    return float(dimension)*FACTOR


def hash_vector(v, tolerance=TOLERANCE):
    """Hashes a Vector by rounding coordinates to the given tolerance.

    This method allows to use vector as keys in dictionaries, by rounding
    the coordinates to a given tolerance.

    Args:
        v (FreeCAD.Vector): The vector to hash.
        tolerance (float): The tolerance within which vectors should be considered identical.

    Returns:
        int: A hash value for the vector.
    """
    # Convert tolerance to decimal places
    ndigits = int(-1 * (math.log10(tolerance)))
    return hash((round(v.x, ndigits), round(v.y, ndigits), round(v.z, ndigits)))


def color_sh2fc(hexcode:str):
    """Transforms a SweetHome3D color to a FreeCAD color

    Args:
        hexcode (str): The SweetHome color code

    Returns:
        tuple: The FreeCAD color
    """
    # We might have transparency as the first 2 digit
    if isinstance(hexcode, list) or isinstance(hexcode, tuple):
        return hexcode
    if not isinstance(hexcode, str):
        assert False, "Invalid type when calling color_sh2fc(), was expecting a string. Got "+str(hexcode)
    offset = 0 if len(hexcode) == 6 else 2
    return (
        int(hexcode[offset:offset+2], 16),   # Red
        int(hexcode[offset+2:offset+4], 16), # Green
        int(hexcode[offset+4:offset+6], 16), # Blue
        int(hexcode[0:offset], 16)           # ALPHA
        )

def color_fc2sh(color):
    """Transforms a FreeCAD color to a SweetHome3D color

    The FreeCAD color can either be a tuple of float (R,G,B,A)
    or an integer (for instance when retrieving a color parameter)
    In case of a tuple each float is normailized to 255 rounded and
    transformed into its hexadecimal representation. In case of an
    int it is simple transformed into an hexadecimal string.

    Args:
        color (any): The FreeCAD color code

    Returns:
        str: The SweetHome3D color
    """
    if isinstance(color, list) or isinstance(color, tuple):
        hexcode = ''.join(f"{round(255*f):02X}" for f in color)
    elif isinstance(color, int):
        hexcode = f"{color:08X}"
    else:
        assert False, "Invalid type when calling color_fc2sh(), was expecting a tuple or an int. Got "+str(color)
    return ''.join([hexcode[6:8], hexcode[0:6]])

def percent_fc2sh(percent):
    # percent goes from 0 -> 1 in SH3d and 0 -> 100 in FC
    return float(percent)/100.0

def percent_sh2fc(percent):
    # percent goes from 0 -> 1 in SH3d and 0 -> 100 in FC
    return int(float(percent)*100)

def set_fc_property(obj, property_type:str, name:str, description:str, value, default_value=None, valid_values=None, group="SweetHome3D"):
    """Set the attribute of the given object as an FC property

    Note that the method has a default behavior when the value is not specified.

    Args:
        obj (object): The FC object to add a property to
        property_type (str): the FC type of property to add
        name (str): the name of the property to add
        description (str): a short description of the property to add
        value (xml.etree.ElementTree.Element|str): The property's value. Defaults to None.
        valid_values (list): an optional list of valid values
    """
    sanitized_name = name.replace('.', '__')
    _add_fc_property(obj, property_type, sanitized_name, description, group)
    if valid_values:
        setattr(obj, sanitized_name, valid_values)
    if value is None:
        _log(f"Setting obj.{sanitized_name}=None")
        return
    if type(value) is ET.Element or type(value) is type(dict()):
        if property_type == "App::PropertyString":
            default_value = "" if default_value is None else str(default_value)
            value = str(value.get(name, default_value))
        elif property_type == "App::PropertyFloat":
            default_value = 0.0 if default_value is None else float(default_value)
            value = float(value.get(name, default_value))
        elif property_type == "App::PropertyInteger":
            default_value = 0 if default_value is None else int(default_value)
            value = int(value.get(name, default_value))
        elif property_type == "App::PropertyPercent":
            default_value = 0.0 if default_value is None else float(default_value)
            value = percent_sh2fc(value.get(name, default_value))
        elif property_type == "App::PropertyBool":
            default_value = "true" if default_value is None else str(bool(default_value)).lower()
            value = value.get(name, default_value) == "true"
        elif property_type == "App::PropertyAngle":
            default_value = 0 if default_value is None else float(default_value)
            value = f"{math.degrees(float(value.get(name, default_value)))} deg"
        elif property_type == "App::PropertyLength":
            default_value = 0.0 if default_value is None else float(default_value)
            value = f"{value.get(name, default_value)} cm"
        elif property_type == "App::PropertyColor":
            value = color_sh2fc(str(value.get(name, default_value)))
    _log(f"Setting @{obj}.{sanitized_name} = {value}")
    setattr(obj, sanitized_name, value)

def set_sh_attribute(attributes, obj, property, value=None):
    """Set the attribute from the value found in the SweetHome3D group.

    If the value does not exists the attribute is not set.

    Args:
        attributes (dict): the dictionary of attributes to set
        obj (Arch): the Arch object
        property (str): the property to retrieve
        value (any, optional): The preferred value. Defaults to None.
            If the value is not found the value found in the SweetHome3D
            group is used.
    """
    property_type = get_fc_property_type(obj, property)
    if not value and property in obj.PropertiesList:
        # Not value was passed, we use the value from the SweetHome3D group
        value = get_fc_property(obj, property, value)

    if property_type == "App::PropertyString":
        pass
    elif property_type == "App::PropertyFloat":
        pass
    elif property_type == "App::PropertyInteger":
        pass
    elif property_type == "App::PropertyPercent":
        pass
    elif property_type == "App::PropertyBool":
        value = "true" if value else "false"
    elif property_type == "App::PropertyLength":
        value = value.Value
    elif property_type == "App::PropertyAngle":
        value = math.radians(value.Value)
    elif property_type == "App::PropertyColor":
        value = color_fc2sh(value)
    else:
        _wrn(f"Property type {property_type} not supported for {property}")
        pass
    if value:
        attributes[property] = str(value)

def get_fc_property(obj, property, default_value=None):
    """Get the property from the FC object.

    Args:
        obj (object): The FC object to get the property from
        property (str): The property to retrieve
        default_value (any, optional): The default value. Defaults to None.

    Returns:
        any: the value of the property.
    """
    value = default_value
    if property in obj.PropertiesList:
        value = obj.getPropertyByName(property)
    return value

def get_fc_property_type(obj, property):
    """Get the property type from the FC object.

    Args:
        obj (FCObject): the FC object
        property (srt): the property name

    Returns:
        str: The type of the property
    """
    if property in obj.PropertiesList:
        return obj.getTypeIdOfProperty(property)
    return None

def _add_fc_property(obj, property_type, name, description, group="SweetHome3D"):
    """Add an property to the FC object.

    All properties will be added under the 'SweetHome3D' group

    Args:
        obj (object): TheFC object to add a property to
        property_type (str): the FC type of property to add
        name (str): the name of the property to add
        description (str): a short description of the property to add
    """
    if name not in obj.PropertiesList:
        obj.addProperty(property_type, name, group, description)


__all__ = ["ang_sh2fc", "coord_fc2sh", "coord_sh2fc", "dim_fc2sh", "dim_sh2fc", "hash_vector", "set_fc_property"]
