To create a zip without DS_Store and other macos files use:

zip -r HA_Gerber.zip . -x "__MACOSX" "*.DS_Store"


from the folder with all gerbers and drl files (without folders)