#!/bin/bash -ex

trap "echo -e '\x1b[01;31mFailed\x1b[0m'" ERR

compare_expected_to_pptx ()
{
    local -r expected="$1"
    local -r pptx="$2"

    local -r tmp_compare='.tmp_compare'

    rm -rf "$tmp_compare"
    mkdir "$tmp_compare"

    unzip "$pptx" -d "$tmp_compare"
    for slide in "$tmp_compare/ppt/slides"/*.xml
    do
        base_slide=$(basename "$slide")
        diff --unified "$expected/$base_slide" "$slide"
    done

    rm -rf "$tmp_compare"
}

readonly tmp='.tmp'

# Test for expected slide output.
for input in test/*.rst
do
    base=$(basename "$input" '.rst')
    rm -rf "$tmp"
    mkdir "$tmp"
    ./rst2pptx.py "$input" "$tmp/output.pptx"
    compare_expected_to_pptx "test/$base" "$tmp/output.pptx"
    rm -rf "$tmp"
done

# Test alternate input/output method.
rm -rf "$tmp"
mkdir "$tmp"
readonly input='test/bullets.rst'
./rst2pptx.py < "$input" > "$tmp/output.pptx"
compare_expected_to_pptx 'test/bullets' "$tmp/output.pptx"
rm -rf "$tmp"

echo -e '\x1b[01;32mOkay\x1b[0m'
