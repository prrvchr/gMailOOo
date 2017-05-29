#!/bin/sh

usage()
{
    echo "usage: MailmergeUpdate.sh [[[-s source ] [-t target ]] | [-h]]"
}
source=
target=
while [ "$1" != "" ]; do
    case "$1" in
        -s | --source )         shift
                                source="$1"
                                ;;
        -t | --target )         shift
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
