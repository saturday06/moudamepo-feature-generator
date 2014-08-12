#!/bin/sh
# -*- coding: utf-8-unix -*-

cd `dirname $0`
./build.py
cd ..
rm -fr deploy
git clone . deploy
cd deploy
git remote set-url origin git@github.com:saturday06/moudamepo-feature-generator.git
git fetch
git checkout deploy/unix/latest
cp ../README.md ../generatefeature.py .
