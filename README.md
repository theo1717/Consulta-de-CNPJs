# Consulta-de-CNPJs

Olá,

Esse projeto consiste em um código Python que conecta à uma API para consultar informações sobre determinados CNPJs.

# Como funciona

Ao executar o código, os CNPJs no arquivo Excel 'cnpjs.xlsx' são coletados para serem consultados pela API, retornando várias 
informações sobre cada CNPJ coletado.

# Como utilizar

Para tornar a experiência mais prática, sem a necessitade de utilizar git, baixe a pasta .zip do projeto.
Baixe o Visual Studio Code, caso não tenha intalado:

Link para instalação: [baixar](https://code.visualstudio.com/Download)

Após descompactar a pasta, abra ela no Visual Studio Code.
O Excel 'cnpjs.xlsx' já vem com um CNPJ aleatório para demonstração.

Execute o código Python 'consultaEmpresas.py'

Após alguns segundos (se for apenas 1 CNPJ), o programa se encerra e gera um Excel 'resultado_cnpjs.xlsx' no mesmo diretório
do projeto.
Abra o arquivo 'resultado_cnpjs.xlsx' e confira os diversos dados da empresa coletada!

Caso deseje coletar informações sobre mais CNPJs, batsa inserir o(s) CNPJ(s) desejado(s) no arquivo 'cnpjs.xlsx',
executar 'consultaEmpresas.py' e aguardar! O arquivo 'resultado_cnpjs.xlsx' será atualizado após a execução.

# Observação

Para fins demonstrativos, a parte que conecta ao MySQL foi comentada no código, para testar apenas com Excel. Porém sinta-se a vontade
para conectar ao banco caso desejar.