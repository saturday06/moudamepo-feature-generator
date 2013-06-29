#!/bin/sh
# -*- coding: utf-8-unix -*-

cd `dirname $0`

cd ../build
chmod 755 build.py
./build.py

cd ../
chmod 755 generatefeature.py
rm -fr test/got
./generatefeature.py test/input test/got
diff -ru test/expected test/got
