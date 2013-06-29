#!/usr/bin/env python
# -*- coding: utf-8-unix -*-

import base64
import codecs
import os

i = codecs.open('template.py', 'r', 'utf_8')
js = codecs.open('main.js', 'r', 'shift_jis')
ods = open('dummy.ods', 'rb')
odsBase64 = base64.b64encode(ods.read())
width = 76
odsBase64 = '\n'.join(odsBase64[pos:pos+width] for pos in xrange(0, len(odsBase64), width))
template = i.read()
template = template.replace("%%%GENERATE_FEATURE_JS%%%", js.read())
template = template.replace("%%%GENERATE_FEATURE_DUMMY_ODS%%%", odsBase64)
template = template.replace("\r\n", "\n")
i.close()

target = '../generatefeature.py'
o = codecs.open(target, 'w', 'utf_8')
o.write(template)
o.close()
os.chmod(target, 0755)
