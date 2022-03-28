Se o namespace `CONTOSO` não estiver disponível no menu de preenchimento automático, siga as etapas a seguir para registrar o suplemento no Excel.

### <a name="excel-on-windows-or-mac"></a>[Excel para Windows ou Mac](#tab/excel-windows)

1. No Excel, escolha a guia **Inserir** e, em seguida, escolha a seta para baixo localizada à direita de **Meus Suplementos**.

    :::image type="content" source="../images/select-insert.png" alt-text="Captura de tela da faixa de opções Inserir no Excel no Windows, com a seta para baixo Meus suplementos realçada.":::

1. Na lista de suplementos disponíveis, localize a seção **Suplementos do desenvolvedor** e selecione o seu suplemento **contagem de estrelas** para registrá-lo.

    :::image type="content" source="../images/list-starcount.png" alt-text="Captura de tela da faixa de opções Inserir no Excel no Windows, com o suplemento Funções Personalizadas do Excel destacado na lista Meus suplementos.":::

# <a name="excel-on-the-web"></a>[Excel na Web](#tab/excel-online)

1. No Excel, escolha a guia **Inserir** e, em seguida, escolha **Suplementos**.

    :::image type="content" source="../images/excel-cf-online-register-add-in-1.png" alt-text="Captura de tela da faixa de opções Inserir no Excel na web, com o botão Meus suplementos destacado.":::

1. Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.

1. Escolha **Procurar...** e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.

1. Selecione o arquivo **manifest. XML** e escolha **abrir**, escolha **Carregar**.

1. Agora, vamos experimentar a nova função. Na célula **B1**, digite o texto **=CONTOSO. GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")** e pressione Enter. Você deve ver que o resultado na célula **B1** é o número atual de estrelas fornecido para o [repositório do GitHub de funções personalizadas do Excel](https://github.com/OfficeDev/Excel-Custom-Functions).

---
