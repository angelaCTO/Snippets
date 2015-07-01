#! /bin/bash

# untar all files (dir)
for a in `ls -1 *.tar.gz`; do tar -zxvf $a; done

# unpackage all files (subdir)
find . -type d -exec sh -c '(cd {} && (for a in `ls -l *.rpm`; do sudo rpm -ivh $a; done))' ';'
