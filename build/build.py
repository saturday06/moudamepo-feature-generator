#!/usr/bin/env python
# -*- coding: utf-8-unix -*-

import base64
import codecs
import os

i = codecs.open('template.py', 'r', 'utf_8')
js = codecs.open('main.js', 'r', 'shift_jis')
od = open('dummy.odg', 'rb')
odBase64 = base64.b64encode(od.read())
width = 76
odBase64 = '\n'.join(odBase64[pos:pos+width] for pos in xrange(0, len(odBase64), width))
template = i.read()
template = template.replace("%%%GENERATE_FEATURE_JS%%%", js.read())
template = template.replace("%%%GENERATE_FEATURE_DUMMY_OD%%%", odBase64)
template = template.replace("\r\n", "\n")
i.close()

target = '../generatefeature.py'
o = codecs.open(target, 'w', 'utf_8')
o.write(template)
o.close()
os.chmod(target, 0755)
