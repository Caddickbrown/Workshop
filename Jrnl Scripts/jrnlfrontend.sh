PS3='Option: '
options=("Write an Entry" "Show All of Today's Entries" "Show a Random Entry" "Tacos" "Info" "Quit")
printf '\033[8;40;100t' # Set the size of the window
printf 'Welcome to the Jrnl Frontend, below are your options.\n\nWhat would you like to do?.\n\n'
select selection in "${options[@]}"; do
    case $selection in
        "Write an Entry")
            read dataentry
            jrnl $dataentry
            printf "Sorted!\n\n"
            ;;
        "Show All of Today's Entries")
            printf '\nHere are todays entries... \n\n'
            jrnl -on today
            ;;
        "Show a Random Entry")
            printf '\nHere is a random entry... \n\n'
            jrnl -on "$(jrnl --short | shuf -n 1 | cut -d' ' -f1,2)"
            ;;
        "Tacos")
            printf "\nAccording to NationalTacoDay.com, Americans are eating 4.5 billion $selection each year.\n\n"
            ;;
        "Info")
            printf "\nThis frontend was created to make working with Jrnl a little easier. Commonly used commands can be accessed quickly and easily rather than typing them out each time.\n\n" | fold -s
            ;;
	    "Quit")
	        exit
	        ;;
        *) printf "\nInvalid option $REPLY\n\n";;
    esac
    printf 'Can I help you with anything else?\n\n'
done