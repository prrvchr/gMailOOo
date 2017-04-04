#!/bin/sh

usage()
{
    echo "usage: MailmergeUpdate.sh [[[-s sourcefile ] [-t targetfile ]] | [-h]]"
}
source=
target=
while [ "$1" != "" ]; do
    case "$1" in
        -s | --sourcefile )     shift
                                source="$1"
                                ;;
        -t | --targetfile )     shift
                                target="$1"
                                ;;
        -h | --help )           usage
                                exit
                                ;;
        * )                     usage
                                exit 1
    esac
    shift
done
sudo mv "$target" "$target".bak && sudo cp "$source" "$target"
