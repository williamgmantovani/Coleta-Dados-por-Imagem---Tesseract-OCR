# Coleta_dados_OCR
Este Script realiza a leitura da imagem, recorte da imagem, tratamento da imagem, leitura por Tesseract OCR e armazenamento em planilha.

A programação desenvolvida, chamada a partir daqui de script, utiliza uma biblioteca chamada Pytesseract e o software Tesseract OCR, que realiza a identificação de caracteres de vários formatos através de reconhecimento óptico. No caso, foi coletado informações da imagem e transformados em caractere e armazenados em uma tabela.

O script é iniciado com a importação das bibliotecas necessárias para a sua aplicação. Neste início é inicializado um cronômetro para acompanhamento o tempo de processamento de cada imagem e acessar o caminho (PATH) do software Tesseract OCR, na leitura das imagens.

Assim, são definidas as funções “coleta_area” e “tratamento_imagem”, que são apresentados na sequência.

O script, em sua execução, realiza o seguinte procedimento: acessa uma pasta com todas as imagens capturadas e abre a imagem individualmente. Com a imagem acessada, foi definido em tuplas a área na imagem, em pixels, onde está cada informação a ser coletada. Por exemplo, a data na imagem é definida pela tupla (1798, 8, 1914, 28) e significa que, na imagem a informação data está no pixel no eixo x de 1798 a 1914 pixels e no eixo y de 8 a 28 pixels, formando assim uma área da data. Com essas tuplas definidas para cada variável a coletar, a imagem acessa a função “coleta_area” e realiza o corte da imagem para dada informação a ser coletada.

Esta imagem recortada na sua área da imagem, acessa a função “tratamento_imagem” para que seja realizado o tratamento da imagem e facilitar a interpretação do Tesseract OCR para o caractere desejado com o menor processamento.

Esses tratamentos de imagens são na ordem: (i) o redimensionamento da imagem (zoom), (ii) o desfoque gaussiano para focar no texto, (iii) a redução de ruídos na imagem, zerado os canais de RED e BLUE da imagem, (iv) conversão da imagem para escala de cinza e a (v) binarização da imagem, onde as intensidades de cor de 0 a 127 são convertidos em 0 (cor preta) e as intensidade de cor de 128 a 255 são convertidos em 1 (cor branca).

Com este tratamento, a imagem tratada é enviada ao software Tessaract OCR para a realizar a melhor conversão em caractere da área definida. Este ciclo é realizado para cada área de cada imagem que possui a informação necessária para a coleta.

Após a descoberta da área da imagem em texto, cada texto é salvo em uma planilha, como apresentado na amostra da Figura 5. Cada linha apresenta uma imagem e cada coluna apresenta a informação coletada de cada imagem, que no total foram de 26 para este estudo.
