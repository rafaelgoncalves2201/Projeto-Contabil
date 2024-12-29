# Automação Contábil com Python

Este projeto tem como objetivo automatizar o preenchimento de dados contábeis em uma plataforma web. Ele utiliza ferramentas como Selenium para automação do navegador e docx para extração de dados de arquivos do Microsoft Word (.docx).

## Funcionalidades

- Login automático em uma plataforma web contábil.
- Extração de dados de balanços patrimoniais a partir de arquivos Word.
- Preenchimento automático de formulários na plataforma web.
- Processamento de múltiplos relatórios de forma sequencial.

## Tecnologias Utilizadas

- **Python**: Linguagem principal do projeto.
- **Selenium**: Biblioteca para automação de navegadores.
- **python-docx**: Biblioteca para manipulação de documentos Word.
- **Time**: Gerenciamento de pausas durante a execução do código.
- **OS**: Manipulação de diretórios e arquivos.

## Requisitos

- Python 3.7 ou superior [Baixe aqui](https://www.python.org/downloads/).
- Navegador Google Chrome instalado.
- Driver do Chrome (chromedriver) compatível com a versão do navegador.
- Biblioteca Selenium: `pip install selenium`
- Biblioteca python-docx: `pip install python-docx`

## Instalação

1. Clone este repositório:

   ```bash
   git clone https://github.com/rafaelgoncalves2201/Projeto-Contabil.git
   ```

2. Instale as dependências:

   ```bash
   pip install -r requirements.txt
   ```

3. Certifique-se de que o `chromedriver` está configurado no PATH ou no diretório do projeto.

## Uso

1. Certifique-se de que o arquivo `chromedriver` está no diretório correto.
2. Altere os caminhos dos arquivos e diretórios no código conforme necessário.
3. Execute o script:

   ```bash
   python automacao_contabil.py
   ```

4. O script irá:
   - Realizar login na plataforma web contábil.
   - Processar os relatórios em formato .docx dentro da pasta configurada.
   - Preencher automaticamente os campos do formulário e enviar os dados.

## Estrutura do Projeto

```
/
|-- automacao_contabil.py  # Script principal
|-- /relatorios            # Pasta contendo arquivos .docx com os relatórios
```

## Personalização

- **Dados de login**: Altere os valores de email e senha diretamente no código ou implemente um arquivo de configuração.
- **Caminho dos relatórios**: Atualize o caminho para a pasta de relatórios no código.
- **Seletores XPATH**: Certifique-se de que os XPATHs correspondem à estrutura atual da plataforma web.

## Licença

Este projeto está sob a licença MIT. Consulte o arquivo LICENSE para mais informações.

