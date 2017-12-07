#!/bin/bash

# Just a simple script that calls the main python script to generate the test
# documentation
PYTHON='python3'

function execute {
    echo "Input: $INPUT, Output: $OUTPUT"
    $PYTHON $(dirname $0)/../src/doc_gen/doc_gen.py -i $INPUT -o $OUTPUT
    echo
}

script_dir=$(dirname $0)
tests_dir="${script_dir}/../src/tests"
out_dir="${script_dir}/../out"
mkdir ${out_dir}

echo "Simple 1-to-1 generation from Markdown"
INPUT="${tests_dir}/single"
OUTPUT="${out_dir}/single_gen.docx"
execute

echo "Appending to end of template"
INPUT="${tests_dir}/single"
OUTPUT="${out_dir}/appending.docx"
execute

echo "Placing inline for several sections"
INPUT="${tests_dir}/multi"
OUTPUT="${out_dir}/multi.docx"
execute
