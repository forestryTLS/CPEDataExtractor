GREENC="\033[32m"
ENDC="\033[0m"

TARGET_BRANCH="pre-enrolment-table"
CURRENT_BRANCH="$(git rev-parse --abbrev-ref HEAD)"

if [ $CURRENT_BRANCH != $TARGET_BRANCH ]; then
    echo -e "Switching to target branch \"$TARGET_BRANCH\"...\n"
    git checkout $TARGET_BRANCH > /dev/null 2>&1
fi

echo -e "Activating virtual enviornment...\n"
source venv/Scripts/activate

echo -e "Running script...\n"
python get_data.py --status Active Completed Concluded

echo -e "Deactivating virtual environment...\n"
deactivate

# revert to the previously selected branch
if [ $CURRENT_BRANCH != $TARGET_BRANCH ]; then
    echo -e "Switching back to \"$CURRENT_BRANCH\" branch...\n"
    git checkout $CURRENT_BRANCH > /dev/null 2>&1
fi

echo -e "${GREENC}Script execution complete.${ENDC}"