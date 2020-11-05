import sys
import re
from restrictiondiction import rsites
from snapgene_reader import snapgene_file_to_dict

x=sys.argv[1:]
commandfiles=len(x)
notthere=[]
shared=[]

for file in x:
    file=snapgene_file_to_dict(file)
    file=file['seq']
    totfile=file.upper()
    for site in rsites:
        search=rsites[site].replace("/","")
        there=re.search(search,totfile)
        if not there:
            search2="[^ACGT\/]"
            badsite=re.search(search2,rsites[site])
            if not badsite:
                notthere.append(site)

occurancedictionary={}
for site in notthere:
    occurance=notthere.count(site)
    occurancedictionary[site]=occurance
for site in occurancedictionary:
    if occurancedictionary[site]==commandfiles:
        shared.append(site)

for site in shared:
    print(site,rsites[site])
