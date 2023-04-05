#!/bin/sh
set -e

if [ -z "$3" ]; then
  echo USAGE: inpptx json outpptx
  exit
fi

touch "$3"

IN_PPTX="$(realpath "$1")"
IN_JSON="$(realpath "$2")"
OUT_PPTX="$(realpath "$3")"
NUMERIC="$(realpath numeric.json)"

if [ -f "$IN_PPTX" -a -f "$IN_JSON" -a -f "$OUT_PPTX" ]; then
  true
else
  exit
fi

TMP=$(mktemp -d)

cd "$TMP"
unzip -q "$IN_PPTX"

# unzip embeddings
if [ -d ppt/embeddings ]; then
  for f in ppt/embeddings/*; do
    mkdir $f.dir
    unzip -qd $f.dir $f
    rm $f
  done
fi

# format all xml files
for f in $(find . -name '*.xml'); do
  xmllint --format --output $f.fmt $f
  rm $f
  mv $f.fmt $f
done

# do substitutions
for ks in $(jq -r keys[] "$IN_JSON"); do

  # remove placeholder leftovers
  grep -lr "$ks" | xargs sed -i '' "s~[_%{]</~</~"
  grep -lr "$ks" | xargs sed -i '' "s~>[}%_]~>~"
  
  # string keys
  v="$(jq -r .$ks "$IN_JSON")"
  grep -lr "$ks" | xargs sed -i '' "s~[_%{]$ks[}%_]~$v~"
  grep -lr "$ks" | xargs sed -i '' "s~>$ks</~>$v</~"
  
  # numeric keys
  kn=$(jq -r .$ks "$NUMERIC")
  if [ "x$kn" = xnull ]; then continue; fi
  kn="9${kn}9"
  grep -lr ">$kn<" | xargs sed -i '' "s~>$kn</~>$v</~"
done

for kn in $(jq -r .[] "$NUMERIC"); do
  kn="9${kn}9"
  grep -lr ">$kn<" | xargs sed -i '' "s~>$kn</~>0</~"
done

# repack file
if [ -d ppt/embeddings ]; then
  cd ppt/embeddings
  for f in *.dir; do
    cd $f
    zip -Drq ../${f%%.dir} .
    cd ..
    rm -r $f
  done
  cd ../..
fi

rm "$OUT_PPTX"
zip -Drq "$OUT_PPTX" .

