import customtkinter
import time

import win32com.client as win32


janela = customtkinter.CTk()
janela.geometry("400x550+100+10") # Largura x Altura + distancia da esquerda + distancia da direita
janela.title("Projeto e-mails")
#janela.iconbitmap("icon.ico")

def def_botao_entrada():
    hora_atual = time.localtime()
    tempo2 = time.strftime("%Y/%m/%d, %H:%M", hora_atual)
    #print("hora atual:", tempo2)
    
    with open("Registo.txt","a", encoding = "utf-8") as registo_arq:       
        registo_arq.write("Hora de entrada: " + tempo2 + "\n")
    
    janela_registo_confirmado = customtkinter.CTkToplevel(janela)
    janela_registo_confirmado.title("Projeto e-mails ")
    janela_registo_confirmado.geometry("300x150+500+100")

    registo_confirmado_label = customtkinter.CTkLabel(janela_registo_confirmado, text = "Registo Confirmado", font = customtkinter.CTkFont(size=16, weight = "bold"))
    registo_confirmado_label.pack(padx=10, pady=(20,10))

    botao_voltar = customtkinter.CTkButton(janela_registo_confirmado, text = " Voltar ", corner_radius=8, command = janela_registo_confirmado.destroy)
    botao_voltar.pack(padx=10, pady=10)
    pass

def def_botao_pausa():
    hora_atual = time.localtime()
    tempo2 = time.strftime("%Y/%m/%d, %H:%M", hora_atual)
    #print("hora atual:", tempo2)
    
    with open("Registo.txt","a", encoding = "utf-8") as registo_arq:       
        registo_arq.write("Pausa: " + tempo2 + "\n")

    janela_registo_confirmado = customtkinter.CTkToplevel(janela)
    janela_registo_confirmado.title("Projeto e-mails ")
    janela_registo_confirmado.geometry("300x150+500+100")

    registo_confirmado_label = customtkinter.CTkLabel(janela_registo_confirmado, text = "Registo Confirmado", font = customtkinter.CTkFont(size=16, weight = "bold"))
    registo_confirmado_label.pack(padx=10, pady=(20,10))

    botao_voltar = customtkinter.CTkButton(janela_registo_confirmado, text = " Voltar ", corner_radius=8, command = janela_registo_confirmado.destroy)
    botao_voltar.pack(padx=10, pady=10)
    pass

def def_botao_saida():
    hora_atual = time.localtime()
    tempo2 = time.strftime("%Y/%m/%d, %H:%M", hora_atual)
    #print("hora atual:", tempo2)
    
    with open("Registo.txt","a", encoding = "utf-8") as registo_arq:       
        registo_arq.write("Hora de saida: " + tempo2 + "\n")

    janela_registo_confirmado = customtkinter.CTkToplevel(janela)
    janela_registo_confirmado.title("Projeto e-mails ")
    janela_registo_confirmado.geometry("300x150+500+100")

    registo_confirmado_label = customtkinter.CTkLabel(janela_registo_confirmado, text = "Registo Confirmado", font = customtkinter.CTkFont(size=16, weight = "bold"))
    registo_confirmado_label.pack(padx=10, pady=(20,10))

    botao_voltar = customtkinter.CTkButton(janela_registo_confirmado, text = " Voltar ", corner_radius=8, command = janela_registo_confirmado.destroy)
    botao_voltar.pack(padx=10, pady=10)

    pass

def def_botao_comentario():
    with open("Registo.txt","a", encoding = "utf-8") as registo2_arq:       
        registo2_arq = registo2_arq.write(comentario_entry.get() + "\n")
    #janela de confirmação
    janela_registo_confirmado = customtkinter.CTkToplevel(janela)
    janela_registo_confirmado.title("Projeto e-mails ")
    janela_registo_confirmado.geometry("300x150+500+100")

    registo_confirmado_label = customtkinter.CTkLabel(janela_registo_confirmado, text = "Registo Confirmado", font = customtkinter.CTkFont(size=16, weight = "bold"))
    registo_confirmado_label.pack(padx=10, pady=(20,10))

    botao_voltar = customtkinter.CTkButton(janela_registo_confirmado, text = " Voltar ", corner_radius=8, command = janela_registo_confirmado.destroy)
    botao_voltar.pack(padx=10, pady=10)
    
    pass

def def_botao_outras_opcoes():
    def def_enviar_dados():
        #with open("Registo.txt","r", encoding = "utf-8") as registo_arq:
        #   registo_arq = registo_arq.read()
        with open("Email_destino.txt", "r", encoding = "utf-8") as email_arq:
            email_enviar = email_arq.read()

        #variáveis em txt dentro do email
        with open("Assunto.txt", "r", encoding = "utf-8") as assunto_arq:
            assunto = assunto_arq.read()
        with open("Corpo_email.txt", "r", encoding = "utf-8") as corpo_email_arq:
            corpo_email = corpo_email_arq.read()

        outlook = win32.Dispatch('outlook.application') #ligar com o outlook
        #criar um email
        email = outlook.CreateItem(0) #associar o remetente
        #cofigurar o email
        email.To = email_enviar
        email.Subject = f"{assunto}"
        email.HTMLBody = f"{corpo_email}"

        #abrir a localização dos anexos
        #with open("Registo.txt", "r", encoding = "utf-8") as anexo1_arq:
        #    anexo1_arq = anexo1_arq.read()

        #print(anexo1_arq)
        
        anexo1_arq = "c:/Users/Caceiro/Desktop/trabalho/projeto_registo_de_horas/custom_tkinter/Registo.txt"
        email.Attachments.Add(anexo1_arq)
        
        email.Send()

        print("enviado")

        pass

    def def_alterar_email():
        #ir buscar o que foi escrito na entry
        with open("Email_destino.txt","w", encoding = "utf-8") as email_destino_arq:       
            email = email_destino_arq.write(alterar_email_entry.get())
        #janela de confirmação
        janela_alterar_email_teste = customtkinter.CTkToplevel(janela_outras_opcoes)
        janela_alterar_email_teste.title("Projeto e-mails ")
        janela_alterar_email_teste.geometry("300x150+900+100")

        alterar_n_emails_label = customtkinter.CTkLabel(janela_alterar_email_teste, text = "E-mail alterado", font = customtkinter.CTkFont(size=16, weight = "bold"))
        alterar_n_emails_label.pack(padx=10, pady=(20,10))

        botao_voltar = customtkinter.CTkButton(janela_alterar_email_teste, text = " Voltar ", corner_radius=8, command = janela_alterar_email_teste.destroy)
        botao_voltar.pack(padx=10, pady=10)
        pass

    #criar nova tela
    janela_outras_opcoes = customtkinter.CTkToplevel(janela)
    janela_outras_opcoes.title("Outras opcões")
    janela_outras_opcoes.geometry("400x550+500+10")
    #caixa de texto    
    registos_anteriores_label = customtkinter.CTkLabel(master = janela_outras_opcoes, text = "Registos anteriores", font = customtkinter.CTkFont(size=16, weight = "bold"))
    registos_anteriores_label.pack(padx=10, pady=(20,10))
    
    #with open("Email_destino.txt","w", encoding = "utf-8") as email_destino_arq:       
    #    email_destino_arq = email_destino_arq.write(comentario_entry.get() + "\n")

    with open("Registo.txt","r", encoding = "utf-8") as registo_arq:
        registo_arq = registo_arq.read()
    
    lista_registos_frame = customtkinter.CTkScrollableFrame(janela_outras_opcoes, width=300, height=270)
    lista_registos_frame.pack(padx=10, pady=10)
    #mostrar só os 60 primeiros porque se pedir 61 vai aparecer umas reticencias no meio
    scrolabletext= customtkinter.CTkLabel(lista_registos_frame, text = registo_arq)
    scrolabletext.pack(padx=10, pady=10)

    botao_enviar_dados = customtkinter.CTkButton(janela_outras_opcoes, text = " Enviar dados ", corner_radius=8, command = def_enviar_dados)
    botao_enviar_dados.pack(padx=10, pady=10)

    with open("Email_destino.txt", "r", encoding = "utf-8") as teste_email_arq:
        teste_email = teste_email_arq.read()
    alterar_email_entry = customtkinter.CTkEntry(janela_outras_opcoes, width=200, placeholder_text= teste_email)
    alterar_email_entry.place(x=40, y=415)
    alterar_email_botao = customtkinter.CTkButton(janela_outras_opcoes, width=80, text = " Alterar ", corner_radius=8, command = def_alterar_email)
    alterar_email_botao.place(x=265, y=415)  

    botao_voltar = customtkinter.CTkButton(janela_outras_opcoes, text = " Voltar ", corner_radius=8, command = janela_outras_opcoes.destroy)
    botao_voltar.pack(padx=10, pady=50)

    pass

# C O M P O N E N T E S   D A   J A N E L A   I N I C I A L

titulo = customtkinter.CTkLabel(master = janela, text = "Registo de Horas", font = customtkinter.CTkFont(size=30, weight = "bold"))
titulo.pack(padx=10, pady=(40,20)) 

botao_entrada = customtkinter.CTkButton(janela, text = " Entrada ", corner_radius=8, command = def_botao_entrada)
botao_entrada.pack(padx=10, pady=10)

botao_pausa = customtkinter.CTkButton(janela, text = " Pausa ", corner_radius=8, command = def_botao_pausa)
botao_pausa.pack(padx=10, pady=10)

botao_saida = customtkinter.CTkButton(janela, text = " Saida ", corner_radius=8, command = def_botao_saida)
botao_saida.pack(padx=10, pady=10)

comentario_entry = customtkinter.CTkEntry(janela, width=150, placeholder_text= "Inserir comentário")
comentario_entry.pack(padx=10, pady=10)
botao_comentario = customtkinter.CTkButton(janela, text = " Guardar comentário ", corner_radius=8, command = def_botao_comentario)
botao_comentario.pack(padx=10, pady=10)

botao_outas_opcoes = customtkinter.CTkButton(janela, text = " Outras opções ", corner_radius=8, command = def_botao_outras_opcoes)
botao_outas_opcoes.pack(padx=10, pady=10)

botao_sair = customtkinter.CTkButton(janela, text = " Sair ", corner_radius=8, command = janela.destroy)
botao_sair.pack(padx=10, pady=10)


# L O O P   P A R A   O   P R O G R A M A   E S T A R   S E M P R E   A   C O R R E R 
janela.mainloop()


"""
print("a")
tempo = time.ctime()
print("hora atual:", tempo)

hora_atual = time.localtime()
tempo2 = time.strftime("%m/%d/%Y, %H:%M:%S", hora_atual)
print("hora atual:", tempo2)
tempo3 = time.strftime("%H:%M", hora_atual)
print("tempo3:", tempo3)


"""