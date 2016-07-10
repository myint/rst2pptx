#!/bin/bash -ex

trap "echo -e '\x1b[01;31mFailed\x1b[0m'" ERR

for input in test/*.rst
do
    base=$(basename "$input" '.rst')
    rm -rf .tmp
    mkdir .tmp
    ./rst2pptx.py "$input" .tmp/output.pptx
    unzip .tmp/output.pptx -d .tmp
    for slide in .tmp/ppt/slides/*.xml
    do
        base_slide=$(basename "$slide")
        diff --unified "test/$base/$base_slide" "$slide"
    done
    rm -rf .tmp
done

echo -e '\x1b[01;32mOkay\x1b[0m'
