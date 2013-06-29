#!/usr/bin/env python
# -*- coding: utf-8-unix -*-

import base64
import codecs
import os

i = codecs.open('template.py', 'r', 'utf_8')
js = codecs.open('main.js', 'r', 'shift_jis')
ods = open('dummy.ods', 'r')
template = i.read().replace("%%%GENERATE_FEATURE_JS%%%", js.read()).replace("%%%GENERATE_FEATURE_DUMMY_ODS%%%", base64.b64encode(ods.read()))
template = template.replace("\r\n", "\n")
i.close()

target = '../generatefeature.py'
o = codecs.open(target, 'w', 'utf_8')
o.write(template)
o.close()
os.chmod(target, 0755)
