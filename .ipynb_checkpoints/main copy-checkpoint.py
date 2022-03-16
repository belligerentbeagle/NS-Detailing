whoandwhenpresent = { #[name:day_of_mount_gone]
    'Aaron': [1,2],
}

names = ["tom","dikc","harry","Aaron"]

for name in names:
    if name in list(whoandwhenpresent.keys()):
        for daynumber in whoandwhenpresent[name]:
            
            print(daynumber)
        