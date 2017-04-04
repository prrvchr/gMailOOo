#!/bin/bash

cd ./gMailOOo/

zip -0 gMailOOo.zip mimetype

zip -r gMailOOo.zip *

cd ..

mv ./gMailOOo/gMailOOo.zip ./gMailOOo.oxt
