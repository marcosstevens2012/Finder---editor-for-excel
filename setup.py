
# -*- coding: utf-8 -*-
from distutils.core import setup
import py2exe

setup(console=['Principal(Richard).py'])

#Ahora bien un ejemplo más completo sería:



import sys
from distutils.core import setup

kwargs = {}
if 'py2exe' in sys.argv:
    import py2exe
    kwargs = {
        'console' : [{
            'script'         : 'Principal(Richard).py',
            'description'    : 'Descripcion del programa.',
            'icon_resources' : [(0, 'icon.ico')]
            }],
        'zipfile' : None,
        'options' : { 'py2exe' : {
            'dll_excludes'   : ['w9xpopen.exe'],
            'bundle_files'   : 1,
            'compressed'     : True,
            'optimize'       : 2
            }},
         }

setup(
    name='Buscador',
    author='Nombre del autor',
    author_email='autor@correo.com',
    **kwargs)