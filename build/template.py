#!/usr/bin/env python
# -*- coding: utf-8-unix; mode: javascript -*-
# vim: set ft=javascript
#
# Automatic feature generation
#
# Usage: 
#   ./generatefeature.py inputdirectory outputdirectory
#

import sys

generateFeatureJs = r""" //" // magic comment for editor's syntax highlighing
%%%GENERATE_FEATURE_JS%%%
"""
#" /* magic comment for editor's syntax highlighting

# http://python3porting.com/noconv.html
if sys.version < '3':
    import codecs
    generateFeatureJs = codecs.raw_unicode_escape_decode(generateFeatureJs)[0]
