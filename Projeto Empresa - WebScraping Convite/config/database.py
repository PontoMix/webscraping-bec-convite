###Banco de Dados com as informações primordiais das variáveis de ambiente###

#Biblioteca para interagir com o SO, aumentar a segurança e poder criar e utilizar as variáveis de ambiente (.env)
import os

#Importando a biblioteca .env para poder utilizá-la no programa e 
#acessar o site sem expor o usuário e senha para fora do ambiente criado com o programa (maior segurança)
from dotenv import load_dotenv

#Função que carregará as variáveis de ambiente.env antes de iniciar a próxima função
load_dotenv()

#Dicionário com as informações essenciais e sensíveis que precisamos para acessar a BEC
database_infos = {
    "login" : os.getenv('LOGIN'),
    "password" : os.getenv('PASSWORD'),
    "username_pc" : os.getenv('USERNAME_PC')
}