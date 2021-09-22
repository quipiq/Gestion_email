import imaplib
import email
import os
from os import walk
import json
import sys
import keyboard
import time

#########remettre laa fonction imprimer sans commentaire


def delete(server, user, password):
    m = connect(server, user, password)
    m.select('Inbox')
    typ, data = m.search(None, 'ALL')
    for num in data[0].split():
        m.store(num, '+FLAGS', '\\Deleted')
    m.expunge()

def graphique_lancement():
    os.system('cls')
    print("Bienvenue dans le gestionnaire d'e-mail                                                              ©Jean-Lou Saulnier\n\n")
    print("Ce programme va imprimer directement sur l'imprimantes en première position\n\n")

def continuer():
    continuer = input("Voulez vous lancer le programme y/n : ")
    if continuer == "y":
        pass
    else:
        sys.exit()

def preset():
    graphique_lancement()
    continuer()
    lance = False
    try:
        with open('data.json') as json_data:
            data = json.load(json_data)
        server = data['Conf_1']['server']
        user = data['Conf_1']['mail']
        password = data['Conf_1']['mdp']
        dossier = data['Conf_1']['dossier']
    except:
        lancement()
    downloadAllAttachmentsInInbox(server, user, password, dossier)
    delete(server, user, password)

def imprimer(dossier):
    listeFichiers = []
    #for (repertoire, sousRepertoires, fichiers) in walk(dossier):
        #listeFichiers.extend(fichiers)
    #os.startfile(f'{dossier}\\{listeFichiers[-1]}', "print")


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
            imprimer(dossier)


def downloadAllAttachmentsInInbox(server, user, password, dossier):
    print("\n\nEn cours de recherche...")
    m = connect(server, user, password)
    resp, items = m.search( None, '(UNSEEN)')
    items = items[0].split()
    for emailid in items:
        if keyboard.is_pressed('q'):  # if key 'q' is pressed 
            sys.exit()
        downloadAttachmentsInEmail(m, emailid.decode("utf-8"), dossier)
        print("\nDocuments télécharger\n")

def lancement():
    graphique_lancement()
    continuer()

    server = input("Veuillez entrer le nom du server imap : ")
    user = input("Veuillez entrer l'e-mail' : ")
    password = input("Veuillez entrer le mot de passe : ")
    dossier = input("Veuillez entrer le nom de dossier de sortie : ")

    new_data = {
        'Conf_1':{
            'server': server,
            'mail': user, 
            'mdp': password, 
            'dossier': dossier
        }
    }
    file = open('data.json', "w")
    json.dump(new_data, file)
    file.close

    detach_dir = '.'
    if dossier not in os.listdir(detach_dir):
        print("salut")
        os.mkdir(dossier)
    downloadAllAttachmentsInInbox(server, user, password, dossier)

preset()

while True:
    time.sleep(180)
    downloadAllAttachmentsInInbox(server, user, password, dossier)
    delete(server, user, password)

    