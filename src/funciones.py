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

# import uno
import unohelper
from es.atareao.libreoffice.Funciones import XFunciones
import hashlib
import base64


unidades = []
decenas = []
centenas = []
unidades.append('cero')
unidades.append('uno')
unidades.append('dos')
unidades.append('tres')
unidades.append('cuatro')
unidades.append('cinco')
unidades.append('seis')
unidades.append('siete')
unidades.append('ocho')
unidades.append('nueve')
unidades.append('diez')
unidades.append('once')
unidades.append('doce')
unidades.append('trece')
unidades.append('catorce')
unidades.append('quince')
unidades.append('dieciseis')
unidades.append('diecisiete')
unidades.append('dieciocho')
unidades.append('diecinueve')
unidades.append('veinte')
unidades.append('veintiuno')
unidades.append('veintidos')
unidades.append('veintitres')
unidades.append('veinticuatro')
unidades.append('veinticinco')
unidades.append('veintiseis')
unidades.append('veintisiete')
unidades.append('veintiocho')
unidades.append('veintinueve')
unidades.append('treinta')
decenas.append('')
decenas.append('')
decenas.append('')
decenas.append('')
decenas.append('cuarenta')
decenas.append('cincuenta')
decenas.append('sesenta')
decenas.append('setenta')
decenas.append('ochenta')
decenas.append('noventa')
centenas.append('')
centenas.append('ciento')
centenas.append('doscientos')
centenas.append('trescientos')
centenas.append('cuatrocientos')
centenas.append('quinientos')
centenas.append('seiscientos')
centenas.append('setecientos')
centenas.append('ochocientos')
centenas.append('novecientos')

unidadesf = []
decenasf = []
centenasf = []
unidadesf.append('cero')
unidadesf.append('una')
unidadesf.append('dos')
unidadesf.append('tres')
unidadesf.append('cuatro')
unidadesf.append('cinco')
unidadesf.append('seis')
unidadesf.append('siete')
unidadesf.append('ocho')
unidadesf.append('nueve')
unidadesf.append('diez')
unidadesf.append('once')
unidadesf.append('doce')
unidadesf.append('trece')
unidadesf.append('catorce')
unidadesf.append('quince')
unidadesf.append('dieciseis')
unidadesf.append('diecisiete')
unidadesf.append('dieciocho')
unidadesf.append('diecinueve')
unidadesf.append('veinte')
unidadesf.append('veintiuna')
unidadesf.append('veintidos')
unidadesf.append('veintitres')
unidadesf.append('veinticuatro')
unidadesf.append('veinticinco')
unidadesf.append('veintiseis')
unidadesf.append('veintisiete')
unidadesf.append('veintiocho')
unidadesf.append('veintinueve')
unidadesf.append('treinta')
decenasf.append('')
decenasf.append('')
decenasf.append('')
decenasf.append('')
decenasf.append('cuarenta')
decenasf.append('cincuenta')
decenasf.append('sesenta')
decenasf.append('setenta')
decenasf.append('ochenta')
decenasf.append('noventa')
centenasf.append('')
centenasf.append('ciento')
centenasf.append('doscientas')
centenasf.append('trescientas')
centenasf.append('cuatrocientas')
centenasf.append('quinientas')
centenasf.append('seiscientas')
centenasf.append('setecientas')
centenasf.append('ochocientas')
centenasf.append('novecientas')


class FuncionesImpl(unohelper.Base, XFunciones):

    def __init__(self, ctx):
        self.ctx = ctx

    def is_integer(self, number):
        try:
            int(number)
            return True
        except Exception as e:
            int(e)
            return False
        return False

    def doobiemult(self, a, b):
        return a * b

    def substring(self, astring, start, length):
        if self.is_integer(start) and self.is_integer(length):
            if start == 1:
                start = 0
                end = length
            else:
                start = start - 1
                end = start + length
            return astring[start:end]
        return ''

    def interpola(self, x1, y1, x2, y2, x):
        a = (x1 * y2 - x2 * y1) / (x1 - x2)
        b = (y1 - y2) / (x1 - x2)
        return a + b * x

    def leenumero(self, number):
        if not self.is_integer(number):
            return ''
        intue = int(number)
        number = str(number)
        longit = len(number)
        if longit < 3:
            if intue < 31:
                numero_leido = unidades[intue]
            elif intue >= 31 and intue <= 39:
                numero_leido = unidades[30] + ' y ' + unidades[intue - 30]
            elif intue == 40:
                numero_leido = decenas[4]
            elif intue >= 41 and intue <= 49:
                numero_leido = decenas[4] + ' y ' + unidades[intue - 40]
            elif intue == 50:
                numero_leido = decenas[5]
            elif intue >= 51 and intue <= 59:
                numero_leido = decenas[5] + ' y ' + unidades[intue - 50]
            elif intue == 60:
                numero_leido = decenas[6]
            elif intue >= 61 and intue <= 69:
                numero_leido = decenas[6] + ' y ' + unidades[intue - 60]
            elif intue == 70:
                numero_leido = decenas[7]
            elif intue >= 71 and intue <= 79:
                numero_leido = decenas[7] + ' y ' + unidades[intue - 70]
            elif intue == 80:
                numero_leido = decenas[8]
            elif intue >= 81 and intue <= 89:
                numero_leido = decenas[8] + ' y ' + unidades[intue - 80]
            elif intue == 90:
                numero_leido = decenas[9]
            elif intue >= 91 and intue <= 99:
                numero_leido = decenas[9] + ' y ' + unidades[intue - 90]
        elif longit < 4:
            parte_izquierda = number[:1]
            parte_derecha = number[-2:]
            if parte_derecha == '00' and parte_izquierda == '1':
                numero_leido = 'cien'
            else:
                numero_leido = (centenas[int(parte_izquierda)] + ' ' +
                                self.leenumero(parte_derecha))
        elif longit < 7:
            parte_izquierda = number[:longit - 3]
            parte_derecha = number[-3:]
            if int(parte_izquierda) == 0:
                numero_leido = self.leenumero(parte_derecha)
            elif int(parte_izquierda) == 1:
                numero_leido = 'mil ' + self.leenumero(parte_derecha)
            elif int(parte_izquierda) > 1:
                numero_leido = (self.leenumero(parte_izquierda) + ' mil ' +
                                self.leenumero(parte_derecha))
        elif longit < 13:
            parte_izquierda = number[:longit - 6]
            parte_derecha = number[-6:]
            if int(parte_izquierda) == 0:
                numero_leido = self.leenumero(parte_derecha)
            elif int(parte_izquierda) == 1:
                numero_leido = 'un millón ' + self.leenumero(parte_derecha)
            elif int(parte_izquierda) > 1:
                numero_leido = (self.leenumero(parte_izquierda) +
                                ' millones ' + self.leenumero(parte_derecha))
        elif longit < 25:
            parte_izquierda = number[:longit - 12]
            parte_derecha = number[-12:]
            if int(parte_izquierda) == 0:
                numero_leido = self.leenumero(parte_derecha)
            elif int(parte_izquierda) == 1:
                numero_leido = 'un billón ' + self.leenumero(parte_derecha)
            elif int(parte_izquierda) > 1:
                numero_leido = (self.leenumero(parte_izquierda) +
                                ' billones ' + self.leenumero(parte_derecha))
        elif longit < 49:
            parte_izquierda = number[:longit - 24]
            parte_derecha = number[-24:]
            if int(parte_izquierda) == 0:
                numero_leido = self.leenumero(parte_derecha)
            elif int(parte_izquierda) == 1:
                numero_leido = 'un trillón ' + self.leenumero(parte_derecha)
            elif int(parte_izquierda) > 1:
                numero_leido = (self.leenumero(parte_izquierda) +
                                ' trillones ' + self.leenumero(parte_derecha))
        elif longit < 97:
            parte_izquierda = number[:longit - 48]
            parte_derecha = number[-48:]
            if int(parte_izquierda) == 0:
                numero_leido = self.leenumero(parte_derecha)
            elif int(parte_izquierda) == 1:
                numero_leido = 'un cuatrillón ' + self.leenumero(parte_derecha)
            elif int(parte_izquierda) > 1:
                numero_leido = (self.leenumero(parte_izquierda) +
                                ' cuatrillones ' +
                                self.leenumero(parte_derecha))
        elif longit < 193:
            parte_izquierda = number[:longit - 96]
            parte_derecha = number[-96:]
            if int(parte_izquierda) == 0:
                numero_leido = self.leenumero(parte_derecha)
            elif int(parte_izquierda) == 1:
                numero_leido = 'un quintillón ' + self.leenumero(parte_derecha)
            elif int(parte_izquierda) > 1:
                numero_leido = (self.leenumero(parte_izquierda) +
                                ' quintillones ' +
                                self.leenumero(parte_derecha))
        return numero_leido.upper()[:1] + numero_leido.lower()[1:]

    def lee_numerof(self, number):
        if not self.is_integer(number):
            return ''
        intue = int(number)
        number = str(number)
        longit = len(number)
        # trios = longit / 3
        if longit < 3:
            if intue < 31:
                numero_leido = unidadesf[intue]
            elif intue >= 31 and intue <= 39:
                numero_leido = unidadesf[30] + ' y ' + unidadesf[intue - 30]
            elif intue == 40:
                numero_leido = decenasf[4]
            elif intue >= 41 and intue <= 49:
                numero_leido = decenasf[4] + ' y ' + unidadesf[intue - 40]
            elif intue == 50:
                numero_leido = decenasf[5]
            elif intue >= 51 and intue <= 59:
                numero_leido = decenasf[5] + ' y ' + unidadesf[intue - 50]
            elif intue == 60:
                numero_leido = decenasf[6]
            elif intue >= 61 and intue <= 69:
                numero_leido = decenasf[6] + ' y ' + unidadesf[intue - 60]
            elif intue == 70:
                numero_leido = decenasf[7]
            elif intue >= 71 and intue <= 79:
                numero_leido = decenasf[7] + ' y ' + unidadesf[intue - 70]
            elif intue == 80:
                numero_leido = decenasf[8]
            elif intue >= 81 and intue <= 89:
                numero_leido = decenasf[8] + ' y ' + unidadesf[intue - 80]
            elif intue == 90:
                numero_leido = decenasf[9]
            elif intue >= 91 and intue <= 99:
                numero_leido = decenasf[9] + ' y ' + unidadesf[intue - 90]
        elif longit < 4:
            parte_izquierda = number[:1]
            parte_derecha = number[-2:]
            if parte_derecha == '00' and parte_izquierda == '1':
                numero_leido = 'cien'
            else:
                numero_leido = (centenasf[int(parte_izquierda)] + ' ' +
                                self.lee_numerof(parte_derecha))
        elif longit < 7:
            parte_izquierda = number[:longit - 3]
            parte_derecha = number[-3:]
            if int(parte_izquierda) == 0:
                numero_leido = self.lee_numerof(parte_derecha)
            elif int(parte_izquierda) == 1:
                numero_leido = 'mil ' + self.lee_numerof(parte_derecha)
            elif int(parte_izquierda) > 1:
                numero_leido = (self.lee_numerof(parte_izquierda) + ' mil ' +
                                self.lee_numerof(parte_derecha))
        elif longit < 13:
            parte_izquierda = number[:longit - 6]
            parte_derecha = number[-6:]
            if int(parte_izquierda) == 0:
                numero_leido = self.lee_numerof(parte_derecha)
            elif int(parte_izquierda) == 1:
                numero_leido = 'un millón ' + self.lee_numerof(parte_derecha)
            elif int(parte_izquierda) > 1:
                numero_leido = (self.lee_numerof(parte_izquierda) +
                                ' millones ' +
                                self.lee_numerof(parte_derecha))
        elif longit < 25:
            parte_izquierda = number[:longit - 12]
            parte_derecha = number[-12:]
            if int(parte_izquierda) == 0:
                numero_leido = self.lee_numerof(parte_derecha)
            elif int(parte_izquierda) == 1:
                numero_leido = 'un billón ' + self.lee_numerof(parte_derecha)
            elif int(parte_izquierda) > 1:
                numero_leido = (self.lee_numerof(parte_izquierda) +
                                ' billones ' +
                                self.lee_numerof(parte_derecha))
        elif longit < 49:
            parte_izquierda = number[:longit - 24]
            parte_derecha = number[-24:]
            if int(parte_izquierda) == 0:
                numero_leido = self.lee_numerof(parte_derecha)
            elif int(parte_izquierda) == 1:
                numero_leido = 'un trillón ' + self.lee_numerof(parte_derecha)
            elif int(parte_izquierda) > 1:
                numero_leido = (self.lee_numerof(parte_izquierda) +
                                ' trillones ' +
                                self.lee_numerof(parte_derecha))
        elif longit < 97:
            parte_izquierda = number[:longit - 48]
            parte_derecha = number[-48:]
            if int(parte_izquierda) == 0:
                numero_leido = self.lee_numerof(parte_derecha)
            elif int(parte_izquierda) == 1:
                numero_leido = ('un cuatrillón ' +
                                self.lee_numerof(parte_derecha))
            elif int(parte_izquierda) > 1:
                numero_leido = (self.lee_numerof(parte_izquierda) +
                                ' cuatrillones ' +
                                self.lee_numerof(parte_derecha))
        elif longit < 193:
            parte_izquierda = number[:longit - 96]
            parte_derecha = number[-96:]
            if int(parte_izquierda) == 0:
                numero_leido = self.lee_numerof(parte_derecha)
            elif int(parte_izquierda) == 1:
                numero_leido = ('un quintillón ' +
                                self.lee_numerof(parte_derecha))
            elif int(parte_izquierda) > 1:
                    numero_leido = (self.lee_numerof(parte_izquierda) +
                                    ' quintillones ' +
                                    self.lee_numerof(parte_derecha))
        return numero_leido.upper()[:1] + numero_leido.lower()[1:]

    def capitaliza(self, text):
        return text.capitalize()

    def tituliza(self, text):
        return text.title()

    def encodebase64(self, text):
        return base64.b64encode(text.encode())

    def decodebase64(self, text):
        return base64.b64decode(text).decode()

    def sha1(self, text):
        m = hashlib.sha1()
        m.update(text.encode())
        return m.hexdigest()

    def sha256(self, text):
        m = hashlib.sha1()
        m.update(text.encode())
        return m.hexdigest()

    def md5(self, text):
        m = hashlib.md5()
        m.update(text.encode())
        return m.hexdigest()


def createInstance(ctx):
    return FuncionesImpl(ctx)


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    createInstance,
    "es.atareao.libreoffice.Funciones.python.FuncionesImpl",
    ("com.sun.star.sheet.AddIn",),)

if __name__ == '__main__':
    fi = FuncionesImpl(None)
    print(fi.leenumero(111223))
    print(fi.tituliza('esto es una prueba'))