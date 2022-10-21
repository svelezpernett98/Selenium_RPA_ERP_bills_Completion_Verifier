from pickle import TRUE
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from openpyxl import load_workbook
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time, base64, os
from datetime import datetime, date, timedelta
import win32com.client as win32
from twilio.rest import Client 

class Facture(object):
    def __init__(self):
        pass
    @staticmethod
    def sendArtmodeB2b():

        #Ruta de chromedricer.exe
        dir = os.getcwd()
        PATH = str(dir)+"/chromedriver.exe"

        #declaramos variable de webdriver
        init = webdriver.Chrome(PATH)

        #voy a saya
        init.get("website url")

        #declaro variables de login input, password input, insertamos credenciales y click en ingresar
        login_button = WebDriverWait(init, 60).until(EC.presence_of_element_located((By.ID, 'uname')))
        login_button.send_keys("username")
        password_button = WebDriverWait(init, 60).until(EC.presence_of_element_located((By.ID, 'pwd'))) 
        password_button.send_keys("password")
        enter_button = WebDriverWait(init, 60).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[1]/div/div/table/tbody/tr[5]/td[2]/input'))).click()

        #declaro variable de barra de busqueda, inserto "b2b", enter y click en la opcion requerida del menu
        search_bar = WebDriverWait(init, 60).until(EC.presence_of_element_located((By.ID, 'buscarcontrol')))
        search_bar.send_keys("b2b")
        search_bar.send_keys(Keys.RETURN)
        side_menu_button = WebDriverWait(init, 60).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[12]/div[1]/ul/li[4]/a'))).click()

        #acepto la alerta de chrome
        time.sleep(2)
        init.switch_to.alert.accept()
        time.sleep(2)

        #cambio a iframe
        init.switch_to.frame(init.find_element_by_xpath('//*[@id="inferior"]'))
        time.sleep(3)

        #ciclo para tomar el valor del estado de factura, y validar si es pendiente, en cuyo caso toma el tiempo. Si pasaron mas de 3 hrs
        #envia un correo y SMS
        n = 0
        status_input = []
        for i in range(1, 6):
            #Toma el valor del ESTADO
            status_input = init.find_element_by_xpath(f"/html/body/div[4]/div[2]/div/div/div/form/div/div["+str(i)+"]/input[6]").get_attribute('value')
            print(status_input)
            #toma el valor del numero de la factura
            bill_number_input = init.find_element_by_xpath(f"/html/body/div[4]/div[2]/div/div/div/form/div/div["+str(i)+"]/input[1]").get_attribute('value')
            #toma el valor de la clasificaion de la factura
            type_input = init.find_element_by_xpath(f"/html/body/div[4]/div[2]/div/div/div/form/div/div["+str(i)+"]/input[2]").get_attribute('value')
        
            if status_input != "FINALIZADO":
                #Toma el valor de la hora final
                final_time_input = init.find_element_by_xpath(f"/html/body/div[4]/div[2]/div/div/div/form/div/div["+str(i)+"]/input[5]").get_attribute('value')

                #compara la hora actual con la hora final de la factura, si pasaron mas de 10800 segundos(3 hr), envia un correo
                def substract_hour(hour1,hour2):
                    #Resta la hora actual con la hora final del input, para obtener la cantidad de hroas que han pasado desde que la hora final
                    time_format = "%H:%M:%S"
                    h1 = datetime.strptime(hour1, time_format)
                    h2 = datetime.strptime(hour2, time_format)
                    time_difference = h1 - h2
                    total_seconds = time_difference.total_seconds()
                    hours = int(total_seconds / 60 / 60)
                    total_seconds -= hours*60*60
                    minutes = int(total_seconds/60)
                    total_seconds -= minutes*60
                    
                    #Si han pasado mas de 3 horas desde la hora final del input:
                    if hours > 3:
                        #Cambia la hora en saya (Adelanta 10 minutos de la hora actual, en multiplos del 10)
                        #Toma individualmente la hora, los minutos y si es AM o PM (en reloj de 12 hrs)
                        minute_now = str(now.strftime("%M"))
                        hour_now = str(now.strftime("%I"))
                        am_pm_now = str(now.strftime("%p"))
                        
                        #toma el segundo digito del minuto para despues restarlo del minuto y sumarle 10
                        #por ejemplo si el minuto es 37, toma el 7 y resta 37 - 7 y le suma 10 al resultado que seria 40
                        second_digit_of_minute = minute_now[1]

                        #Toma el input de la hora inicio y su valor
                        initial_time_input = WebDriverWait(init, 60).until(EC.presence_of_element_located((By.XPATH, f"/html/body/div[4]/div[2]/div/div/div/form/div/div["+str(i)+"]/input[4]")))
                        initial_time_input_value = init.find_element_by_xpath(f"/html/body/div[4]/div[2]/div/div/div/form/div/div["+str(i)+"]/input[4]").get_attribute('value')
                        
                        #Presiona la flecha derecha 2 veces en el input de la hora inicio, para seleccionar la hora 
                        initial_time_input.send_keys(Keys.ARROW_LEFT, Keys.ARROW_LEFT)
                        print(initial_time_input_value)
                        
                        #Resta el segundo digito del minuto con el minuto y le suma 10
                        #por ejemplo si el minuto es 37, toma el 7 y resta 37 - 7 y le suma 10 al resultado que seria 40
                        minute0 = int(minute_now) - int(second_digit_of_minute) + 10
                        print(minute0)
                        
                        #Si la operacion del minuto da 60 le indica que el minuto debe ser 00 y que la hora se debe aumentar en 1
                        if minute0 == 60:
                            minute0 = str(0)
                            hour_now = int(hour_now) + 1
                            print(minute0)
                                          
                            #En caso de que la hora sea 12 y se le valla a sumar 1 porque el minuto dio 60, le indica que la hora NO debe ser 13, sino que debe ser 1, ya que el reloj de saya el de 12 hrs
                            if hour_now == 12:
                                hour_now = 1
                        else:
                            print("melo")
                        
                        #construye una hora que sea la hora actual pero con el minuto modificado para que sea el siguiente multiplo del diez 
                        final_time = now.strftime(f"{hour_now}:{minute0}:{am_pm_now}")
                        
                        #Envia la hora que construimos al input de la hora inicio
                        initial_time_input.send_keys(str(final_time))
                        time.sleep(2)
                        
                        #Si la hora actual es "AM", se envia la tecla "a", o "p", dependiendo si la hora actual es pm o am y asi cambiar el input de la hora inicial de am a pm o viceversa
                        if am_pm_now == "AM":
                            initial_time_input.send_keys("a")
                            print("es pm")
                        else:
                            initial_time_input.send_keys("p")
                            print("es am")
                        time.sleep(10)
                        
                        #Envia correo
                        outlook = win32.Dispatch('outlook.application')
                        mail = outlook.CreateItem(0)
                        mail.To = 'test@email.co'
                        mail.Subject = f'Factura en estado "{status_input}" despues de 3 horas de la hora inicial, factura atrasada!'
                        mail.Body = 'Mensaje generado por TI'
                        mail.HTMLBody = (f'<h2>Mensaje generado RPA: factura {bill_number_input} de {type_input} en estado "{status_input}", atrasada por '+str(f"{hours}:{minutes}:{total_seconds}")+' segundos en ' + str(today) + '</h2>') #this field is optional
                        mail.Send()
                        
                        #Envia SMS con twilio 
                        account_sid = 'twilio_Account_sid' 
                        auth_token = 'twilio_auth_token' 
                        client = Client(account_sid, auth_token) 
 
                        message = client.messages.create(  
                                                    messaging_service_sid='messaging_service_sid', 
                                                    body=(f'SMS Generado automaticamente: Mensaje generado RPA: factura {bill_number_input} de {type_input} en estado {status_input} atrasada por '+str(f"{hours}:{minutes}:{total_seconds}")+' segundos en ' + str(today)),      
                                                    to='+57phone#' 
                                                ) 
                        print(message.sid)
                    else:
                        #Envia correo
                        outlook = win32.Dispatch('outlook.application')
                        mail = outlook.CreateItem(0)
                        mail.To = 'test@email.co'
                        mail.Subject = f'Factura en estado "{status_input}" todavia dentro del tiempo establecido'
                        mail.Body = 'Mensaje generado por TI'
                        mail.HTMLBody = (f'<h2>Mensaje generado RPA: factura {bill_number_input} de {type_input} en estado "{status_input}" todavia dentro del tiempo establecido, tiempo: '+str(f"{hours}:{minutes}:{total_seconds}")+' segundos en ' + str(today) + '</h2>') #this field is optional
                        mail.Send()
                        
                        #Envia SMS con twilio 
                        account_sid = 'twilio_account_sid' 
                        auth_token = 'twilio_auth_token' 
                        client = Client(account_sid, auth_token) 
 
                        message = client.messages.create(  
                                                    messaging_service_sid='messaging_service_sid', 
                                                    body=(f'Mensaje generado RPA: factura {bill_number_input} de {type_input} en estado "{status_input}" todavia dentro del tiempo establecido, tiempo: '+str(f"{hours}:{minutes}:{total_seconds}")+' segundos en ' + str(today)),      
                                                    to='+57phone#' 
                                                ) 
                        print(message.sid)

                now = datetime.now()
                today = date.today()
                current_time = now.strftime("%H:%M:%S")
        
                substract_hour(current_time, final_time_input)

            #Si las primeras 3 facturas estan en "FINALIZADO", envia un correo indicandolo
            else:
                n = n+1
                print(n)
                if n >= 5:
                    #Envia correo
                    outlook = win32.Dispatch('outlook.application')
                    mail = outlook.CreateItem(0)
                    mail.To = 'test@email.co'
                    mail.Subject = 'Todas las facturas revisadas estan en estado "FINALIZADO", todo esta funcionando correctamente!'
                    mail.Body = 'Mensaje generado por TI'
                    mail.HTMLBody = '<h2>Mensaje generado RPA: Todo parece estar funcionando correctamente, no hay facturas atrasadas</h2>' #this field is optional
                    mail.Send()
                    
                    #Envia SMS con twilio 
                    account_sid = 'twilio_account_sid' 
                    auth_token = 'twilio_auth_token' 
                    client = Client(account_sid, auth_token) 
 
                    message = client.messages.create(  
                                                messaging_service_sid='messaging_service_sid', 
                                                body='SMS Generado automaticamente: Todas las facturas revisadas estan en estado "FINALIZADO", todo esta funcionando correctamente!',      
                                                to='+57phone#' 
                                            ) 
                    print(message.sid)
                else:
                    print("Nada")

def main():
        Facture.sendArtmodeB2b()

if __name__ == "__main__":
    main()

























