PS3='Option: '
options=("Write an Entry" "Show All of Today's Entries" "Show a Random Entry" "Edit Today's Entries" "Edit All Entries" "Open the folder" "Tacos" "Info" "Config File" "Quit")
printf '\033[8;40;100t' # Set the size of the window
printf 'Welcome to the Jrnl Frontend, below are your options.\n\nWhat would you like to do?.\n\n'
select selection in "${options[@]}"; do
    case $selection in
        "Write an Entry")
            printf '\nWhat do you want to journal about?\n\n'
            read dataentry
            jrnl $dataentry
            printf "Sorted!\n\n"
            ;;
        "Show All of Today's Entries")
            printf '\nHere are todays entries... \n\n'
            jrnl -on today
            ;;
        "Show a Random Entry")
            printf '\nHere is a random entry...\n\n'
            jrnl -on "$(jrnl --short | shuf -n 1 | cut -d' ' -f1,2)"
            ;;
        "Edit Today's Entries")
            printf '\nLoading todays entries...\n\n'
            jrnl -on today --edit
            ;;
        "Edit All Entries")
            printf '\nLoading all entries...\n\n'
            jrnl --edit
            ;;
        "Open the folder")
            printf '\nOpening Jrnl Folder...\n\n'
            output=$(jrnl --list)
            folder_path=$(echo "$output" | awk -F '->' '/default/{sub(/^[ \t]+/, "", $2); gsub(/[ \t]+$/, "", $2); print $2}')
            explorer "$folder_path"
            ;;
        "Tacos")
            printf "\nAccording to NationalTacoDay.com, Americans are eating 4.5 billion $selection each year.\n\n"
            ;;
        "Info")
            printf "\nThis frontend was created to make working with Jrnl a little easier. Commonly used commands can be accessed quickly and easily rather than typing them out each time.\n\n" | fold -s
            ;;
        "Config File")
            printf '\nSoon you will be able to access the config file from here... but not now... sorry!\n\n'
            output=$(jrnl --list)
            config_path=$(echo "$output" | awk 'NR==1{print $NF}')
            config_path=$(echo "$config_path" | sed 's/^\(.*\)$/\1/' | tr -d '()')
            start "$config_path"
            ;;
	    "Quit")
	        exit
	        ;;
        *) 
            printf "\nInvalid option $REPLY\n\n";;
    esac
    printf 'Can I help you with anything else?\n\n'
done
