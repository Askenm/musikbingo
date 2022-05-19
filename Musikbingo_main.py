import email
import streamlit as st
import streamlit.components.v1 as components
from spotipy.oauth2 import SpotifyClientCredentials
import spotipy
import random
import pandas as pd
from docx import Document
from docx.enum.section import WD_ORIENT
import os
from zipfile import ZipFile
import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
import smtplib, ssl
import shutil


def send_mail(send_from, send_to, subject, text, files=None,
              server="127.0.0.2"):
    assert isinstance(send_to, list)

    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(text))

    for f in files or []:
        with open(f, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(f)
            )
        # After the file is closed
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
        msg.attach(part)

    context = ssl.create_default_context()
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.ehlo()  # Can be omitted
        server.starttls(context=context)
        server.ehlo()  # Can be omitted
        server.login(send_from, 'rgvqjvalatxnpokm')
        server.sendmail(send_from, send_to, msg.as_string())
        server.close()

st.header('Musikbingo')

playlist_link = st.text_input('Spotify Playlist Link:')
email_address = st.text_input('Modtager Email:')
num_plader = st.slider("antal bingoplader")




cid = '898e9fabd3b54c05ad85d793598616d5'
secret  = 'bb9c762c66734a4493393b2df70297f0'
client_credentials_manager = SpotifyClientCredentials(client_id=cid, client_secret=secret)
sp = spotipy.Spotify(client_credentials_manager = client_credentials_manager)

#playlist_uri = '37i9dQZF1DX0kbJZpiYdZl'

#uri_link = 'https://open.spotify.com/embed/playlist/37i9dQZF1DX0kbJZpiYdZl' + playlist_uri

if 'lav_bingo' not in st.session_state:
    st.session_state['lav_bingo'] = False

if st.button('Lav Bingoplader!'):
    st.session_state['lav_bingo'] = True

if st.session_state['lav_bingo']:
    playlist_URI = playlist_link.split("/")[-1].split("?")[0]
    track_uris = [x["track"]["uri"] for x in sp.playlist_tracks(playlist_URI)["items"]]


    value_strings = [track["track"]["name"]+'\n'+track["track"]["artists"][0]["name"]
                    for track in sp.playlist_tracks(playlist_URI)["items"]]


    ixes = [i for i in range(len(value_strings))]

    max_songs_pr_sheet = 10
    bingoplade_dfs = []

    for plade in range(num_plader):
        chosen_songs = []
        cur_plade = {'kol_{}'.format(i) : [] for i in range(7)}
        print(plade)
        num_songs_pr_row = {i:3 for i in range(3)}
        num_songs_pr_row[random.choice([0,1,2])]+=1

        for row in range(3):
            blank_ixes = [i for i in range(7)]
            song_ixes = [blank_ixes.pop(random.choice([i for i in range(len(blank_ixes))])) for j in range(num_songs_pr_row[row])]

            for col in range(7):
                if col in song_ixes:
                    song = None
                    while song == None or song in chosen_songs:
                        song = random.choice(value_strings)
                        
                    
                    cur_plade['kol_{}'.format(col)].append(song)
                    chosen_songs.append(song)
                else:
                    cur_plade['kol_{}'.format(col)].append('')


        bingoplade_dfs.append(pd.DataFrame.from_dict(cur_plade))

    dir_name = 'bingoplader_'+str(random.random())[2:]
    os.mkdir(dir_name)
    zipObj = ZipFile(dir_name+'/bingoplader.zip', 'w')
    st.info("Laver Bingoplader fra playlisten {}...".format(sp.playlist(playlist_link)['name']))
    for ix,plade_df in enumerate(bingoplade_dfs):
        for col in plade_df.columns:
            plade_df[plade_df[col][0]] = plade_df[col]
            plade_df.drop(col,axis=1,inplace=True)
        plade_df.drop(0,axis=0,inplace=True)
        document = Document()
        style = document.styles['Normal']
        font = style.font
        font.name = 'Calibri'
        section = document.sections[-1]
        new_width, new_height = section.page_height, section.page_width
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = new_width
        section.page_height = new_height

        #document.add_heading('MUSIKBINGO')

        t = document.add_table(plade_df.shape[0]+1, plade_df.shape[1])

        # add the header rows.
        for j in range(plade_df.shape[-1]):
            t.cell(0,j).text = plade_df.columns[j]

        # add the rest of the data frame
        for i in range(plade_df.shape[0]):
            for j in range(plade_df.shape[-1]):
                t.cell(i+1,j).text = str(plade_df.values[i,j])

        document.save(dir_name+'/bingoplade_{}.docx'.format(ix+1))
        zipObj.write(dir_name+'/bingoplade_{}.docx'.format(ix+1))
        
    zipObj.close()
    st.info("Sender bingoplader til "+email_address)
    send_mail('aske.osv@gmail.com',[email_address],'Musikbingo','Go Quiz!',[dir_name+'/bingoplader.zip'])
    shutil.rmtree(dir_name)
    st.session_state['lav_bingo'] = False
    st.balloons()