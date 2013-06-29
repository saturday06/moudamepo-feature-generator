#!/bin/sh
# -*- coding: utf-8-unix -*-

cd `dirname $0`

git checkout master
./build.py
cd ..
cp README.md README.md.backup
mv generatefeature.py generatefeature.py.backup

git checkout deploy/unix/latest

mv README.md.backup README.md
mv generatefeature.py.backup generatefeature.py
