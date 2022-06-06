import yfinance as yf
import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
import urllib.request as ur
import imaplib
import email
from email.header import decode_header
import webbrowser
import os
from st_aggrid import AgGrid, GridOptionsBuilder
from st_aggrid.shared import GridUpdateMode


def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)




st.sidebar.title('Opções')
option = st.sidebar.selectbox('Selecione:', ('','Emails','Analises'))
st.header(option)

if option == '':
    st.title('LeBot')
    st.subheader('Robô-Analista Inteligente de Relatórios Anuais de Companhias')
    # st.image('https://media.istockphoto.com/photos/robot-with-financial-and-technical-data-analysis-graph-showing-stock-picture-id1128852644?s=2048x2048')
    st.image('http://ioop.com.br/lebot.png')


if option == 'Emails':
    st.subheader('Analisando fila de Emails')
    # account credentials
    df = pd.DataFrame()
    username = "lebotfairfax@gmail.com"
    password = "ejpijurvfkpdgena"
    # use your email provider's IMAP server, you can look for your provider's IMAP server on Google
    # or check this page: https://www.systoolsgroup.com/imap/
    # for office 365, it's this:
    imap_server = "imap.gmail.com"
    # create an IMAP4 class with SSL 
    imap = imaplib.IMAP4_SSL(imap_server)
    # authenticate
    imap.login(username, password)

    imap.select('Inbox')

    status, messages = imap.select("INBOX")
    # number of top emails to fetch
    N = 5
    # total number of emails
    messages = int(messages[0])


    for i in range(messages, messages-N, -1):
        # fetch the email message by ID
        res, msg = imap.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                # decode the email subject
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    # if it's a bytes, decode to str
                    subject = subject.decode(encoding)
                # decode email sender
                From, encoding = decode_header(msg.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)
                # st.write("Subject:", subject)
                # st.write("From:", From)
                # if the email message is multipart
                if msg.is_multipart():
                    # iterate over email parts
                    for part in msg.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            # get the email body
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            # print text/plain emails and skip attachments
                            # st.write(body)
                            df2 = {'De': From, 'Assunto': subject, 'Texto': body}
                            df = df.append(df2, ignore_index = True)

                        elif "attachment" in content_disposition:
                            # download attachment
                            filename = part.get_filename()
                            if filename:
                                folder_name = clean(subject)
                                if not os.path.isdir(folder_name):
                                    # make a folder for this email (named after the subject)
                                    os.mkdir(folder_name)
                                filepath = os.path.join(folder_name, filename)
                                # download attachment and save it
                                open(filepath, "wb").write(part.get_payload(decode=True))
                else:
                    # extract content type of email
                    content_type = msg.get_content_type()
                    # get the email body
                    body = msg.get_payload(decode=True).decode()
                    if content_type == "text/plain":
                        # print only text email parts
                         print(body)
                if content_type == "text/html":
                    # if it's HTML, create a new HTML file and open it in browser
                    folder_name = clean(subject)
                    if not os.path.isdir(folder_name):
                        # make a folder for this email (named after the subject)
                        os.mkdir(folder_name)
                    filename = "index.html"
                    filepath = os.path.join(folder_name, filename)
                    # write the file
                    open(filepath, "w").write(body)
                    # open in the default browser
                    webbrowser.open(filepath)
                
                # st.write("="*100)
    # close the connection and logout
    imap.close()
    imap.logout()
    df = df[['De', 'Assunto', 'Texto']]
    # st.table(df)

    def aggrid_interactive_table(df: pd.DataFrame):
        options = GridOptionsBuilder.from_dataframe(
        df, enableRowGroup=True, enableValue=True, enablePivot=True
        )

        options.configure_side_bar()

        options.configure_selection("single")
        selection = AgGrid(
            df,
            enable_enterprise_modules=True,
            gridOptions=options.build(),
            theme="light",
            update_mode=GridUpdateMode.MODEL_CHANGED,
            allow_unsafe_jscode=True,
        )

        return selection

    selection = aggrid_interactive_table(df=df)
    if selection:
        st.write("Email selecionado:")
        st.json(selection["selected_rows"])






if option == 'Analises':
    st.subheader('Relatorios da Empresa')    
    st.subheader('Petrobras S.A')

    def aggrid_interactive_table(df: pd.DataFrame):
        options = GridOptionsBuilder.from_dataframe(
        df, enableRowGroup=True, enableValue=True, enablePivot=True
        )

        options.configure_side_bar()

        options.configure_selection("single")
        selection = AgGrid(
            df,
            enable_enterprise_modules=True,
            gridOptions=options.build(),
            theme="light",
            update_mode=GridUpdateMode.MODEL_CHANGED,
            allow_unsafe_jscode=True,
        )

        return selection

    # We will use Amazon stocks
    symbol = 'PETR4.SA'

        # Get stock data
    get_data = yf.Ticker(symbol)

    # Set the time line of your data
    ticket_df = get_data.history(period='1d', start='2022-1-01', end='2022-05-05')

    st.subheader('Visão Geral')
    st.image('http://ioop.com.br/petr4-1.png') 

    st.subheader('Perfil')
    st.image('http://ioop.com.br/petr4-2.png') 

    st.subheader('Resultados')
    st.image('http://ioop.com.br/petr4-3.png') 
    

    st.subheader('Noticias Internacionais')
    news = pd.DataFrame(get_data.news)
    aggrid_interactive_table(df=news)
    

    st.subheader('Acionistas')
    selection = aggrid_interactive_table(df=get_data.institutional_holders) 

    st.subheader('Calendário de Ganhos')
    aggrid_interactive_table(df=get_data.calendar) 

    st.subheader('Sustentatibilidade')
    aggrid_interactive_table(df=get_data.sustainability) 
    # Show your data in line chart
    st.subheader('Fechamento')
    st.line_chart(ticket_df.Close)
    st.subheader('Volume')
    st.line_chart(ticket_df.Volume)
   
    # index= 'MSFT'
    # url_is = 'https://finance.yahoo.com/quote/' + index + '/financials?p=' + index 
    # url_bs = 'https://finance.yahoo.com/quote/' + index +'/balance-sheet?p=' + index
    # url_cf = 'https://finance.yahoo.com/quote/' + index + '/cash-flow?p='+ index

    # st.write(url_is)
    # read_data = ur.urlopen(url_is).read() 
    # soup_is= BeautifulSoup(read_data,'xml')

    # st.write(read_data) 
    # ls= [] # Create empty list
    # for l in soup_is.find_all('div'): 
    # #Find all data structure that is 'div'
    #     ls.append(l.string) # add each element one by one to the list
    #     ls = [e for e in ls if e not in ('Operating Expenses','Non-recurring Events')] # Exclude those columns

    # new_ls = list(filter(None,ls))
    # new_ls = new_ls[12:]
    # is_data = list(zip(*[iter(new_ls)]*6))

    # Income_st = pd.DataFrame(is_data[0:])
    # Income_st.columns = Income_st.iloc[0] # Name columns to first row of dataframe
    # Income_st = Income_st.iloc[1:,] # start to read 1st row
    # Income_st = Income_st.T # transpose dataframe
    # Income_st.columns = Income_st.iloc[0] #Name columns to first row of dataframe
    # Income_st.drop(Income_st.index[0],inplace=True) #Drop first index row
    # Income_st.index.name = '' # Remove the index name
    # Income_st.rename(index={'ttm': '12/31/2019'},inplace=True) #Rename ttm in index columns to end of the year
    # Income_st = Income_st[Income_st.columns[:-5]] # remove last 5 irrelevant columns

    # st.write(Income_st)




