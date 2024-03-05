printf '\033[8;40;100t'
jrnl -on "$(jrnl --short | shuf -n 1 | cut -d' ' -f1,2)"
read -p "Press enter to continue"