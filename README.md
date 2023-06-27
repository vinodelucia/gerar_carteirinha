# gerar_carteirinha
O código, em resumo, lê informações de uma planilha, cria uma carteirinha para cada aluno com base nessas informações, adiciona o texto e a foto do aluno em uma imagem de fundo e salva as carteirinhas como arquivos JPEG.

O código começa importando as bibliotecas necessárias: PIL para manipulação de imagens, ImageDraw para desenhar na imagem e ImageFont para definir as fontes de texto. Além disso, é importada a biblioteca openpyxl para trabalhar com arquivos do Excel.

A imagem de fundo é aberta utilizando o método Image.open(). O caminho absoluto para a imagem é fornecido como argumento. Essa imagem será utilizada como base para a carteirinha.

Em seguida, são definidas as fontes de texto a serem utilizadas na carteirinha. O arquivo de fonte .ttf é fornecido como argumento para o método ImageFont.truetype(). No exemplo, utiliza-se a fonte Arial com tamanhos de 26 e 18 para o nome/RA e para as datas de emissão/validade, respectivamente.

O caminho absoluto do arquivo da planilha é definido na variável caminho_arquivo_excel.

A planilha é aberta utilizando o método openpyxl.load_workbook(), passando o caminho do arquivo como argumento. Em seguida, a primeira aba da planilha é selecionada com o atributo active.

Um loop é iniciado para percorrer as linhas da planilha. A função iter_rows() é utilizada para iterar pelas linhas da planilha a partir da segunda linha (ignorando o cabeçalho) e values_only=True é utilizado para retornar apenas os valores das células.

Dentro do loop, as informações de cada aluno são extraídas da linha da planilha. Cada valor é atribuído a uma variável correspondente, como nome, RA, data de validade, data de emissão e caminho da foto.

Em seguida, é feita uma tentativa de abrir a foto do aluno utilizando o caminho fornecido na planilha. A função Image.open() é utilizada para abrir a imagem e, se for bem-sucedida, a foto é redimensionada para o tamanho desejado (270x220 pixels). Se houver algum erro ao abrir a imagem, uma mensagem de erro é exibida e o loop continua com a próxima linha.

Um novo objeto de imagem é criado copiando a imagem de fundo original. Isso é feito utilizando o método copy() da imagem de fundo.

Um objeto ImageDraw é criado a partir da nova imagem para permitir o desenho de texto na imagem utilizando o método ImageDraw.Draw().

O texto com as informações do aluno (nome, RA, data de emissão e data de validade) é adicionado à imagem utilizando o método text() do objeto ImageDraw. São especificadas as coordenadas de posicionamento do texto, a string de texto em si, a fonte a ser utilizada e a cor do texto (preto neste caso, representado pelo valor RGB (0, 0, 0)).

A foto do aluno é colada na nova imagem utilizando o método paste() da imagem. As coordenadas de posicionamento da foto são especificadas (60, 350).

A nova imagem é convertida para o modo RGB utilizando o método convert() para garantir que esteja no formato apropriado para salvar como arquivo JPEG.

Por fim, a nova imagem é salva como um arquivo JPEG na pasta de destino especificada, com o nome do aluno como o nome do arquivo.
