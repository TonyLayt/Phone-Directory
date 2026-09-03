# -*- coding: utf-8 -*-
"""
Created on Sat Dec  3 15:02:20 2022

@author: TonyLayt
"""
import socket
import os
import configparser

class Server ():        
    def start_config(self):    
        self.config = configparser.ConfigParser()
        self.ini_config()
        self.ip = str(self.config['Connection Settings']['IP-адреса(TCP/IPv4)'])
        self.port = int(self.config['Connection Settings']['Port'])
        self.main_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.main_socket.bind((self.ip, self.port))
        self.main_socket.listen(5)
           
    def ini_config (self):
        if not self.config.read('Settings_server.ini'):
            self.config['Connection Settings'] = {'IP-адреса(TCP/IPv4)':'localhost', 'Port':'8000'}
            with open ('Settings_server.ini', 'w') as configfile:
                self.config.write(configfile)
             
    def find_client (self):
        print (f'Сервер телефонного довідника [v.0.6]\n(c) TonyLayt\n\nадреса сервера: {self.ip}:{self.port} \n\nОчікує підключення...')
        while True:
            new_socket, addres = self.main_socket.accept()
            
            try:
        
                folder_path = 'docdat'
                            
                file_list = os.listdir(folder_path)
                new_socket.send(str(len(file_list)).encode())  # Отправляем количество файлов клиенту
            
                for file_name in file_list:
                    file_path = os.path.join(folder_path, file_name)
                    file_size = os.path.getsize(file_path)
            
                    # Отправляем имя файла и его размер клиенту
                    new_socket.send(file_name.encode())
                    new_socket.send(str(file_size).encode())
            
                    # Открываем файл и по частям передаем его содержимое клиенту
                    with open(file_path, 'rb') as file:
                        while True:
                            file_data = file.read(1024)
                            if not file_data:
                                break
                            new_socket.send(file_data)
                new_socket.close()
                print (f'{addres} Підключився. Надіслано нові данні')
            except:
                print ('Існуюче з''єднання було примусово закрито')
        

if __name__ == '__main__':        
    open_server = Server()
    open_server.start_config()
    open_server.find_client()

         
        





        



