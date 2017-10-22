#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#
# This file is part of Funciones
#
# Copyright (C) 2016 Lorenzo Carbonell
# lorenzo.carbonell.cerezo@gmail.com
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.

# A unique ID for the add-in.
addin_url = 'es.atareao.libreoffice'
addin_name = 'Funciones'
addin_python_file = 'funciones.py'
excel_addin_name = 'Funciones.xlam'
addin_id = "%s.%s" % (addin_url, addin_name)
addin_version = "0.0.3"
addin_displayname = "Funciones varias"
addin_publisher_link = "http://www.atareao.es"
addin_publisher_name = "Lorenzo Carbonell"
addin_functions = [
    {
        'function_name': 'leenumero',
        'function_description': 'Convierte un numero en texto',
        'parameters': [
            ('number', 'El numero a convertir')]
    },
    {
        'function_name': 'interpola',
        'function_description': 'Interpola linealmente entre dos puntos',
        'parameters': [
            ('x1', 'Coordenada x del primer punto'),
            ('y1', 'Coordenada y del primer punto'),
            ('x2', 'Coordenada x del segundo punto'),
            ('y2', 'Coordenada y del segundo punto'),
            ('x', 'Coordenada x del punto que queremos calcular')]
    },
    {
        'function_name': 'substring',
        'function_description': 'Extrae una cadena de texto de otra',
        'parameters': [
            ('astring', 'Cadena original'),
            ('start', 'Posición del primer carácter de la cadena a extraer'),
            ('length', 'Longitud de la cadena a extraer')]
    },
    {
        'function_name': 'doobiemult',
        'function_description': 'Multiplica dos numeros',
        'parameters': [
            ('a', 'Numero 1'),
            ('b', 'Numero 2')]
    },
    {
        'function_name': 'capitaliza',
        'function_description': 'Cambia la primera letra a mayusculas.',
        'parameters': [
            ('text', 'El texto a capitalizar')]
    },
    {
        'function_name': 'tituliza',
        'function_description': 'Convierte una cadena en Formato Titulo',
        'parameters': [
            ('text', 'El texto a titulizar')]
    },
    {
        'function_name': 'encodebase64',
        'function_description': 'Cifra un texto en base 64',
        'parameters': [
            ('text', 'El texto a cifrar')]
    },
    {
        'function_name': 'decodebase64',
        'function_description': 'Descifra un texto en base 64',
        'parameters': [
            ('text', 'El texto a descifrar')]
    },
    {
        'function_name': 'sha1',
        'function_description': 'Calcula el resumen utilizando SHA1',
        'parameters': [
            ('text', 'El texto cuyo resumen se quiere calcular')]
    },
    {
        'function_name': 'sha256',
        'function_description': 'Calcula el resumen utilizando SHA256',
        'parameters': [
            ('text', 'El texto cuyo resumen se quiere calcular')]
    },
    {
        'function_name': 'md5',
        'function_description': 'Calcula el resumen utilizando MD5',
        'parameters': [
            ('text', 'El texto cuyo resumen se quiere calcular')]
    },
]


# description.xml
#
#

desc_xml = open('description.xml', 'w')

desc_xml.write('<?xml version="1.0" encoding="UTF-8"?>\n')
desc_xml.write('\t<description \
xmlns="http://openoffice.org/extensions/description/2006" \n')
desc_xml.write('xmlns:d="http://openoffice.org/extensions/description/2006" \
\n')
desc_xml.write('xmlns:xlink="http://www.w3.org/1999/xlink"> \n')
desc_xml.write('\t<dependencies> \n')
desc_xml.write('\t\t<OpenOffice.org-minimal-version value="2.4" \
d:name="OpenOffice.org 2.4"/> \n')
desc_xml.write('\t</dependencies> \n')
desc_xml.write('\t<registration>\n')
desc_xml.write('\t\t<simple-license accept-by="admin" \
default-license-id="ID0" suppress-on-update="true">\n')
desc_xml.write('\t\t<license-text xlink:href="resources/license_es.txt" \
lang="es" license-id="ID0"/>\n')
desc_xml.write('\t\t<license-text xlink:href="resources/license_en.txt" \
lang="en"/>\n')
desc_xml.write('\t\t</simple-license>\n')
desc_xml.write('\t</registration>\n')
desc_xml.write('\t<identifier value="%s" /> \n' % (addin_id))
desc_xml.write('\t<version value="%s" />\n' % (addin_version))
desc_xml.write('\t<display-name>\n')
desc_xml.write('\t\t<name lang="en">%s</name>\n' % (addin_displayname))
desc_xml.write('\t\t<name lang="es">%s</name>\n' % (addin_displayname))
desc_xml.write('\t</display-name>\n')
desc_xml.write('\t<publisher>\n')
desc_xml.write('\t\t<name xlink:href="%s" lang="en">%s</name>\n' % (
    addin_publisher_link, addin_publisher_name))
desc_xml.write('\t\t<name xlink:href="%s" lang="es">%s</name>\n' % (
    addin_publisher_link, addin_publisher_name))
desc_xml.write('\t</publisher>\n')
desc_xml.write('\t<icon>\n')
desc_xml.write('\t\t<default xlink:href="icons/icon.png"/>\n')
desc_xml.write('\t</icon>\n')
desc_xml.write('\t<extension-description>\n')
desc_xml.write('\t\t<src xlink:href="description/pkg-description_en.txt" \
lang="en"/>\n')
desc_xml.write('\t\t<src xlink:href="description/pkg-description_es.txt" \
lang="es"/>\n')
desc_xml.write('\t</extension-description>\n')
desc_xml.write('</description>')

desc_xml.close


def add_manifest_entry(xml_file, file_type, file_name):
    xml_file.write('<manifest:file-entry \
manifest:media-type="application/vnd.sun.star.%s" \n' % (file_type))
    xml_file.write('manifest:full-path="%s"/> \n' % (file_name))

# manifest.xml
#
# List of files in package and their types.


manifest_xml = open('manifest.xml', 'w')

manifest_xml.write('<manifest:manifest>\n')
add_manifest_entry(manifest_xml,
                   'uno-typelibrary;type=RDB',
                   'X%s.rdb' % (addin_name))
add_manifest_entry(manifest_xml,
                   'configuration-data',
                   'CalcAddIn.xcu')
add_manifest_entry(manifest_xml,
                   'uno-component;type=Python',
                   addin_python_file)
manifest_xml.write('</manifest:manifest> \n')

manifest_xml.close

# CalcAddIn.xcu
#
#

# instance_id references the named UNO component instantiated by Python code
# (that is my understanding at least).
instance_id = "%s.%s.python.%sImpl" % (addin_url, addin_name, addin_name)

# Name of the corresponding Excel add-in if you want to share documents across
# OOo and Excel.
# excel_addin_name = "DoobieExample.xlam"


def define_function(xml_file, function_name, description, parameters):
    xml_file.write('  <node oor:name="%s" oor:op="replace">\n' % (
        function_name))
    xml_file.write('    <prop oor:name="DisplayName"><value \
xml:lang="en">%s</value></prop>\n' % (function_name))
    xml_file.write('    <prop oor:name="Description"><value \
xml:lang="en">%s</value></prop>\n' % (description))
    xml_file.write('    <prop oor:name="Category"><value>\
Add-In</value></prop>\n')
    xml_file.write('    <prop oor:name="CompatibilityName">\
<value xml:lang="en">AutoAddIn.%s.%s</value></prop>\n' % (
        addin_name, function_name))
    xml_file.write('    <node oor:name="Parameters">\n')

    for p, desc in parameters:
        # Optional parameters will have a displayname enclosed in square
        # brackets.
        p_name = p
        xml_file.write(
            '      <node oor:name="%s" oor:op="replace">\n' % (p_name))
        xml_file.write(
            '        <prop oor:name="DisplayName">\
<value xml:lang="en">%s</value></prop>\n' % (p_name))
        xml_file.write(
            '        <prop oor:name="Description">\
<value xml:lang="en">%s</value></prop>\n' % (desc))
        xml_file.write('      </node>\n')
    xml_file.write('    </node>\n')
    xml_file.write('  </node>\n')


calcaddin_xml = open('CalcAddIn.xcu', 'w')

calcaddin_xml.write('<?xml version="1.0" encoding="UTF-8"?>\n')
calcaddin_xml.write('<oor:component-data \
xmlns:oor="http://openoffice.org/2001/registry" \
xmlns:xs="http://www.w3.org/2001/XMLSchema" \
oor:name="CalcAddIns" oor:package="org.openoffice.Office">\n')
calcaddin_xml.write('<node oor:name="AddInInfo">\n')
calcaddin_xml.write('<node oor:name="%s" oor:op="replace">\n' % (instance_id))
calcaddin_xml.write('<node oor:name="AddInFunctions">\n')

for addin_function in addin_functions:
    define_function(calcaddin_xml,
                    addin_function['function_name'],
                    addin_function['function_description'],
                    addin_function['parameters'])

calcaddin_xml.write('</node>\n')
calcaddin_xml.write('</node>\n')
calcaddin_xml.write('</node>\n')
calcaddin_xml.write('</oor:component-data>\n')

calcaddin_xml.close()

# Done
