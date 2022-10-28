# Leavers_Process365

### Converts a leavers mailbox to shared, removes the licence, asks if you want to:  
1. Remove from GAL  
2. Remove from disitribution lists  
3. Add an auto reply  
4. Add read+manage permissions  
5. Add mailbox forwarding

### How to use
1. Download the .ps1 file
3. Browse to the location you saved the file to, right click the file and select `open with powershell`.

#### You must sign in to 3 different log in boxes. This is due to the three different PS modules that 365 has for manipulating different parts. WIP to migrate this script to use MS Graph, however Graph is missing too many features for this to take up too much of my attention.
