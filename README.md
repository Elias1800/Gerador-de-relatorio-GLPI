# Gerador de Relatório GLPI
![Versão](https://img.shields.io/badge/versão-1.0.0-blue)
![Python](https://img.shields.io/badge/python-3.10%2B-brightgreen)
![GLPI](https://img.shields.io/badge/GLPI-10.x-orange)
![Status](https://img.shields.io/badge/status-Concluído-red)
[![Licença MIT](https://img.shields.io/badge/licença-MIT-yellow)](LICENSE)

Este programa permite a geração de relatórios detalhados no formato Excel a partir de dados extraídos do GLPI. Ele organiza informações como abertura e fechamento de chamados, usuário e outros detalhes relevantes.

## Tabela de Conteúdos
- [Requisitos](#requisitos)
- [Bibliotecas](#bibliotecas)
- [Instalação](#instalação)
- [Configuração do Banco de Dados](#configuração-do-banco-de-dados)
- [Como Usar](#como-usar)
- [Como Compilar para Executável (.exe)](#como-compilar-para-executável-exe)
- [Contribuição](#contribuição)
- [Doações](#doações)
- [Licença](#licença)

## Requisitos
Certifique-se de ter os seguintes itens antes de começar:
- Python 3.10 ou superior.
- Pip configurado.
- Bibliotecas
- Acesso ao banco de dados do GLPI
  - Endereço (host).
  - Usuário e senha.
  - Nome do banco de dados.
    
## Bibliotecas
As bibliotecas usadas no projeto são:

| Biblioteca          | Uso no Projeto                                          |
|---------------------|---------------------------------------------------------|
| **html**            | Manipular entidades HTML em textos.                    |
| **os**              | Gerenciar arquivos e diretórios do sistema.            |
| **re**              | Processar e limpar strings usando expressões regulares.|
| **tkinter**         | Criar interface gráfica para o usuário.                |
| **pandas**          | Manipular dados em tabelas.                            |
| **xlwings**         | Interagir com arquivos Excel.                          |
| **sqlalchemy**      | Conectar e interagir com o banco de dados.             |

Para instalar as bibliotecas necessárias, use o comando abaixo:

  
1. Instalar bibliotecas necessárias:
   
  ```
  pip install pandas xlwings sqlalchemy
  ```

2. Instale o PyInstaller:
   
  ```
  pip install pyinstaller
  ```
As bibliotecas **html**, **os**, **re**, e **tkinter** já fazem parte da biblioteca
padrão do Python e não precisam de instalação adicional.

## Instalação
1. Clone o repositório:
  ```
  git clone https://github.com/Elias1800/Gerador-de-relatorio-GLPI.git
  ```
2. Acesse o diretório do projeto:
  ```
  cd gerador-relatorios-glpi
  ```
## Configuração do Banco de Dados
1. Abra o arquivo Script.py no seu editor de texto ou IDE preferido.
2. Localize o trecho de código responsável pela conexão com o banco de dados:
   ```
    # Função para conectar ao banco de dados usando SQLAlchemy e pymysql
    def conectar_banco():
      try:
          # Criação da URL de conexão com SQLAlchemy
          engine = create_engine('mysql+pymysql://user:senha@endereco-do-database:port/name-database?charset=utf8mb4')
          conn = engine.connect()
          return conn
      except Exception as err:
          print(f"Erro ao conectar: {err}")
          return None
   ```
   
4. Substitua as credenciais de exemplo pelos dados do seu ambiente:
   - `endereco-do-database:port`: Endereço do servidor.
   - `user`: Usuário.
   - `senha`: Senha.
   - `name-database`: Nome do banco de dados.
5. Salve o arquivo após editar.

## Como usar
1. Execute o programa no terminal:
 ```
 python Script.py
 ```

2. Siga as instruções exibidas no terminal para:
   - Selecionar filtros e datas.
   - Gerar os relatórios.
3. Os relatórios serão salvos no formato Excel no diretório especificado ou padrão.

## Como Compilar para Executável (.exe)
Se você deseja compilar o programa para um executável `.exe` no Windows, siga estas etapas:

1. Compile o script com PyInstaller:
  ``` 
  pyinstaller --onefile --noconsole Script.py
  ```

2. O executável será gerado no diretório `dist/`:
  ```
  # Verifique o caminho:
  cd dist
  ```

3. Teste o executável:
  ```
  ./Script.exe
  ```
   
### Observação:
A opção `--noconsole` cria o executável sem abrir o terminal ao executá-lo. Isso é útil para interfaces gráficas ou programas que não precisam de interação no terminal.

## Contribuição

[Veja como contribuir](https://github.com/tiagoporto/.github/blob/main/CONTRIBUTING.md).

## Doações

[![Pix](https://img.shields.io/badge/Pix-Doar-brightgreen)](https://www.gerarpix.com.br/pix?code=tlpz7ucJKhsT8J66PCY4QMjgMgv1FhI_y1AzihkZZUtR6pIvgHYauJVeDaATZSsJr9YN_Fr7gge56mrDYGj5a0smqKwZnkbIkeFdciok1h4rOynFHErOTSK9Eehn2-gFxb8Pvl79HCuhyEDj43M1sEuAcMTZCh45Yiym176CJACyhpMdaEgaiQexnpyZzie9umerRzFhCidaqaGSlD1XZc3JTckhIiqwu63IqIl6tKd93u_ocW_wb_yslzg_Qeq3aZ8)

## **Licença**
Este projeto está licenciado sob a licença [MIT](LICENSE)




