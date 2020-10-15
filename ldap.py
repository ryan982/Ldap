from ldap3 import Server, Connection, ALL
import xlsxwriter

server = Server('ldap.forumsys.com  ', get_info=ALL)
conn = Connection(server, 'cn=read-only-admin,dc=example,dc=com', 'password', auto_bind=True)
#conn.search('uid=euler,dc=example,dc=com','(&(objectclass=top))',attributes=['sn', 'mail','objectClass'])

conn.search('ou=mathematicians,dc=example,dc=com','(&(objectclass=top))', attributes=['uniqueMember'])
#print(conn.entries)
entry = conn.entries[0]
name_mathematicians = entry.uniqueMember
#print(name_mathematicians)
ignore_1 = ',dc=example,dc=com'
ignore_2= 'uid='
i = 0
names = []
for i in range(0, len(name_mathematicians)):
    if ignore_1 and ignore_2 in name_mathematicians[i]:
        name = name_mathematicians[i].replace(ignore_1, "")
        name_1 = name.replace(ignore_2,"")
        names.append(name_1)
print(names)

#chemistry
conn1 = Connection(server, 'cn=read-only-admin,dc=example,dc=com', 'password', auto_bind=True)
conn1.search('ou=chemists,dc=example,dc=com','(&(objectclass=top))', attributes=['uniqueMember'])
entry_chem = conn1.entries[0]
name_chem = entry_chem.uniqueMember
#print(name_chem)
ignore_3 = ',dc=example,dc=com'
ignore_4= 'uid='
i = 0
names_chemist = []
for i in range(0, len(name_chem)):
    if ignore_3 and ignore_4 in name_chem[i]:
        name_chem_1 = name_chem[i].replace(ignore_3, "")
        name_chem_2 = name_chem_1.replace(ignore_4,"")
        names_chemist.append(name_chem_2)
print(names_chemist)


#scientists
conn.search('ou=scientists,dc=example,dc=com','(&(objectclass=top))', attributes=['uniqueMember'])
entry_sci = conn.entries[0]
name_sci = entry_sci.uniqueMember
#print(name_sci)
ignore_5 = ',dc=example,dc=com'
ignore_6= 'uid='
i = 0
names_sci = []
for i in range(0, len(name_sci)):
    if ignore_5 and ignore_6 in name_sci[i]:
        name_sci_1 = name_sci[i].replace(ignore_5, "")
        name_sci_2 = name_sci_1.replace(ignore_6,"")
        names_sci.append(name_sci_2)
print(names_sci)

outworkbook = xlsxwriter.Workbook("letsdoit.xlsx")
outsheet = outworkbook.add_worksheet()

outsheet.write("A1", "Scientists")
outsheet.write("B1", "mathematicians")
outsheet.write("C1", "Chemists")
j = 0

for j in range (0, len(names_chemist)):
        outsheet.write(j+1, 0, names_sci[j])
        outsheet.write(j+1, 1, names[j])
        outsheet.write(j+1,2,names_chemist[j])
outworkbook.close()