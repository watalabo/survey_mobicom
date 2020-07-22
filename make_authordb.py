#!/usr/bin/env python
# -*- coding: utf-8 -*-

import openpyxl
import argparse
import sys
import io
import re

def illegal_char_remover(data):
    ILLEGAL_CHARACTERS_RE = re.compile(
        r'[\000-\010]|[\013-\014]|[\016-\037]|[\x00-\x1f\x7f-\x9f]|[\uffff]')
    """Remove ILLEGAL CHARACTER."""
    if isinstance(data, str):
        return ILLEGAL_CHARACTERS_RE.sub("", data)
    else:
        return data


parser = argparse.ArgumentParser(description="excel hoge hoge")
parser.add_argument("filename", help="filename hog hoge")

args = parser.parse_args()
filename = args.filename

fp = open(filename, "r", encoding="cp932")

text = fp.read()
lines = text.split("\n")

authordb = set()

for line in lines:
#    print(line)
    p = re.compile(r"\([^\)]*?\)")
    line = p.sub("", line)
#    print(line)
    line = line.replace(";", ",")
    line = line.replace(" and ", ",")
    line = line.replace("David E. Culler", "David Culler")
    line = line.replace("  ", " ")
    line = line.replace("  ", " ")
    line = line.replace('"', '')
    line = line.replace('"', '')
    line = line.replace('Kay Roemer', 'Kay Romer')
    line = line.replace('Kaushik R Chowdhury', 'Kaushik Chowdhury')
    line = line.replace('Richard E. Howard', 'Richard Howard')
    
    authors = line.split(",")
    for author in authors:
        author = author.strip()
        author = author.lower()
        if author not in authordb:
            authordb.add(author)

#print(authordb)

authordb = sorted(authordb)

for name in authordb:
    print(name)
