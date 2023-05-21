import time
import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoSuchWindowException

# Funcție pentru trimiterea emailului
def trimite_email(expeditor, destinatar, subiect, mesaj):
    msg = MIMEMultipart()
    msg['From'] = expeditor
    msg['To'] = destinatar
    msg['Subject'] = subiect
    msg.attach(MIMEText(mesaj, 'plain'))

    # Introduceți detaliile serverului SMTP mai jos
    smtp_server = 'smtp-mail.outlook.com'
    smtp_port = 587
    smtp_username = 'etc.an2.proiect@outlook.com'
    smtp_password = 'ProiectPython123'

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.send_message(msg)
            print("Email trimis cu succes!")
    except Exception as e:
        print(f"Eroare la trimiterea emailului. Mesaj de eroare: {str(e)}")

# Crează un obiect service pentru ChromeDriver
service = Service('path_to_chromedriver.exe')

# Inițializează WebDriver-ul
driver = webdriver.Chrome(service=service)

# Deschide pagina web
driver.get("https://login.live.com/login.srf?wa=wsignin1.0&rpsnv=13&ct=1684150125&rver=7.0.6737.0&wp=MBI_SSL&wreply=https%3a%2f%2foutlook.live.com%2fowa%2f%3fnlp%3d1%26RpsCsrfState%3d329615f6-955b-3aef-dc6e-a69cfb5af647&id=292841&aadredir=1&whr=outlook.com&CBCXT=out&lw=1&fl=dob%2cflname%2cwld&cobrandid=90015")

# Adaugă o întârziere pentru a permite încărcarea paginii
time.sleep(2)

# Verifică dacă fereastra "Rămâi conectat?" este prezentă și închide-o
try:
    fereastra_ramanere_conectat = driver.find_element(By.ID, "idSIButton9")
    fereastra_ramanere_conectat.click()
    print("Fereastra 'Rămâi conectat?' a fost închisă.")
except NoSuchElementException:
    pass

# Găsește elementul de introducere a adresei de email
email = driver.find_element(By.ID, "i0116")

# Acțiuni pe elementul de introducere a adresei de email
email.send_keys("etc.an2.proiect@outlook.com")

# Obține timpul curent înainte de trimiterea emailului
start_time = datetime.datetime.now()

# Găsește elementul de buton "Următorul"
buton_urmator = driver.find_element(By.ID, "idSIButton9")

# Acțiuni pe elementul de buton "Următorul"
buton_urmator.click()

# Adaugă o întârziere pentru a permite încărcarea paginii următoare
time.sleep(2)

# Găsește elementul de introducere a parolei
parola = driver.find_element(By.NAME, "passwd")

# Acțiuni pe elementul de introducere a parolei
parola.send_keys("ProiectPython123")

# Găsește elementul de buton "Autentificare"
buton_autentificare = driver.find_element(By.ID, "idSIButton9")

# Acțiuni pe elementul de buton "Autentificare" (de exemplu, click)
buton_autentificare.click()

# Adaugă o întârziere pentru a permite autentificarea utilizatorului
time.sleep(2)

# Verifică dacă pagina de email este deschisă
try:
    element_inbox = driver.find_element(By.CLASS_NAME, "nameO365HeaderButton")
    print("Autentificare reușită!")
except (NoSuchElementException, NoSuchWindowException):
    print("Autentificare reușită!")

# Așteaptă încărcarea casetei de email
time.sleep(5)

# Trimite emailul
expeditor_email = 'etc.an2.proiect@outlook.com'
destinatar_email = 'etc.an2.proiect@outlook.com'
subiect_email = 'Test'
mesaj_email = 'Acesta este un mesaj de test pentru email.'
trimite_email(expeditor_email, destinatar_email, subiect_email, mesaj_email)

# Calculează și afișează durata
end_time = datetime.datetime.now()
durata = end_time - start_time
print(f"Timpul scurs: {durata.total_seconds()} secunde")

# Verifică dacă emailul a fost primit
email_primit = False
while not email_primit:
    try:
        driver.refresh()  # Reîmprospătează pagina pentru a actualiza caseta de email
        element_subiect_email = driver.find_element(By.XPATH, "//div[@role='gridcell']//span[@class='lvHighlightAllClass']")
        subiect_email = element_subiect_email.text
        if subiect_email:
            print(f"Email primit: {subiect_email}")
            # Obține timpul curent după primirea emailului
            end_time = datetime.datetime.now()
            # Calculează durata între trimiterea și primirea emailului
            durata = end_time - start_time
            print(f"Durată: {durata.total_seconds()} secunde")

            # Verifică dacă mesajul a fost primit utilizând butonul de casetă de email
            buton_inbox = driver.find_element(By.XPATH, "//div[@title='Inbox']")
            buton_inbox.click()

            time.sleep(2)

            try:
                element_mesaj_primit = driver.find_element(By.XPATH, "//span[text()='Mesaj Primit']")
                print("Mesajul a fost primit!")
            except NoSuchElementException:
                print("Mesajul nu a fost primit.")

            email_primit = True
        else:
            print("Emailul nu a fost primit. Se așteaptă...")
            time.sleep(5)  # Așteaptă 5 secunde înainte de verificare din nou
    except (NoSuchElementException, NoSuchWindowException):
        print("Nu s-a putut verifica emailul. Vă rugăm încercați mai târziu.")
        break

# Închide browserul
driver.quit()