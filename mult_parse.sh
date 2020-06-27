#!/usr/bin/env bash
folder=$1

while IFS= read -r line; do
  new="${line// /_}"
  if [ "$new" != "$line" ]
    then
      if [ -e "$new" ]
      then
        echo not renaming \""$line"\" because \""$new"\" already exists
      else
        echo moving "$line" to "$new"
      mv "$line" "$new"
    fi
  fi
  full_name=("${new%.*}")
  type="${new##*.}"
  if [ "$type" != "json" ]
    then
    python3 parse_json.py -in "$new" -out "$full_name.json"
  fi
done < <(find "$folder" -type f -not -name '.*')
