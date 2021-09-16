import imaplib
import email
import os
from os import walk

def imprimer():
    mid = "\\"
    listeFichiers = []
    for (repertoire, sousRepertoires, fichiers) in walk(dossier):
        listeFichiers.extend(fichiers)
    while True:
        demande = input("Voulez vous imprimer les documents trouvers y/n : ")
        if demande == "y":
            for i in listeFichiers:
                os.startfile(f'{dossier}{mid}{i}', "print")
            break
        elif demande == "n":
            break
        else:
            print("Veuillez entrer une valeur correct !!\n")


def connect(server, user, password):
    m = imaplib.IMAP4(server)
    m.login(user, password)
    m.select()
    return m


def downloadAttachmentsInEmail(m, emailid, dossier):
    resp, data = m.fetch(emailid, "(BODY.PEEK[])")
    email_body = data[0][1]
    mail = email.message_from_string(email_body.decode("utf-8"))
    if mail.get_content_maintype() != 'multipart':
        return
    for part in mail.walk():
        if part.get_content_maintype() != 'multipart' and part.get('Content-Disposition') is not None:
            open(dossier + '/' + part.get_filename(), 'wb').write(part.get_payload(decode=True))


def downloadAllAttachmentsInInbox(server, user, password, dossier):
    print("\n\nEn cours de recherche...")
    m = connect(server, user, password)
    resp, items = m.search( None, '(UNSEEN)')
    items = items[0].split()
    for emailid in items:
        downloadAttachmentsInEmail(m, emailid.decode("utf-8"), dossier)
        print("\nOpération finis\n")

def lancement():
    detach_dir = '.'
    if dossier not in os.listdir(detach_dir):
        os.mkdir(dossier)
    downloadAllAttachmentsInInbox(server, user, password, dossier)
    imprimer()

os.system('cls')

print("Bienvenue dans le gestionnaire d'e-mail                                                              ©Jean-Lou Saulnier\n\n")

server = input("Veuillez entrer le nom du server imap : ")
user = input("Veuillez entrer l'e-mail' : ")
password = input("Veuillez entrer le mot de passe : ")
dossier = input("Veuillez entrer le nom de dossier de sortie : ")

lancement()