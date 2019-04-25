#!/bin/bash
wget https://www.naturaldocs.org/download/natural_docs/2.0.2/Natural_Docs_2.0.2.zip -O /tmp/ND.zip
unzip /tmp/ND.zip -d /tmp
mkdir $TRAVIS_BUILD_DIR/docs
mkdir $TRAVIS_BUILD_DIR/.ND_Config
cp .nd_project.txt $TRAVIS_BUILD_DIR/.ND_Config/Project.txt
mono /tmp/Natural\ Docs/NaturalDocs.exe $TRAVIS_BUILD_DIR/.ND_Config
python3.5 Excel-Addin-Generator/excelAddinGenerator/main.py $TRAVIS_BUILD_DIR/bin/vbaProject.bin SQLlib.xlam
