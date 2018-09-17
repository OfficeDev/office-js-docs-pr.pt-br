# <a name="build-an-excel-add-in-using-react"></a>Criar um suplemento do Excel usando o React

Neste artigo, você verá um passo a passo do processo de criar um suplemento do Excel usando o React e a API JavaScript do Excel.

## <a name="environment"></a>Ambiente

- **Área de Trabalho do Office**: Verifique se você tem a última versão do Office instalada. Comandos de suplemento precisam da compilação 16.0.6769.0000 ou superior (**16.0.6868.0000** recomendada). Saiba como [Instalar a última versão dos aplicativos do Office](http://aka.ms/latestoffice). 
 
- **Office Online**: Não há configuração adicional. Observe que o suporte para comandos no Office Online para contas de trabalho/escola está em versão prévia.

## <a name="prerequisites"></a>Pré-requisitos

- [Node.js](https://nodejs.org)

- Instale a última versão do [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.
    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-web-app"></a>Criar o aplicativo Web

1. Crie uma pasta na sua unidade local e nomeie-a como **my-addin**. Esse é o local em que você criará os arquivos para seu aplicativo.

2. Navegue até a pasta do seu aplicativo.

    ```bash
    cd my-addin
    ```

3. Use o gerador do Yeoman para gerar o arquivo de manifesto para o seu suplemento. Execute o comando a seguir e responda aos prompts, conforme mostrado na seguinte captura de tela.

    ```bash
    yo office
    ```

    - **Escolha um tipo de projeto:** `Office Add-in project using React framework`
    - **Como deseja nomear seu suplemento?** `My Office Add-in`
    - **Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?:** `Excel`

    ![Gerador do Yeoman](../images/yo-office-excel-react.png)
    
    Depois de concluir o assistente, o gerador criará o projeto e instalará os componentes do nó de suporte.

4.  Abra **src/components/App.tsx**, procure o comentário "Atualizar a cor de preenchimento", altere a cor de preenchimento de "amarelo" para "azul" e salve o arquivo. 

    ```js
    range.format.fill.color = 'blue'

    ```

5. No bloco `return` da função `render` em **src/components/App.tsx**, atualize `<Herolist>` para o código abaixo e salve o arquivo. 

    ```js
      <HeroList message='Discover what My Office Add-in can do for you today!' items={this.state.listItems}>
        <p className='ms-font-l'>Choose the button below to set the color of the selected range to blue. <b>Set color</b>.</p>
        <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={this.click}>Run</Button>
    </HeroList>
    ```

6. Execute as etapas em [Adicionar certificados autoassinados como certificado raiz confiáveis](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) para confiar no certificado do sistema operacional do seu computador de desenvolvimento.

7. Carregue seu suplemento para que ele apareça no Excel. No terminal, execute o comando a seguir: 
    
    ```bash
    npm run sideload
    ```

## <a name="try-it-out"></a>Experimente

1. No terminal, execute o comando a seguir para iniciar o servidor de desenvolvimento.

    Windows:
    ```bash
    npm start
    ```

2. No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-2b.png)

3. Selecione um intervalo de células na planilha.

4. No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como azul.

    ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>Próximas etapas

Você criou com êxito um suplemento do Excel usando o React, parabéns! Agora, saiba mais sobre os recursos dos suplementos do Excel e crie um mais complexo, acompanhando o tutorial de suplemento do Excel.

> [!div class="nextstepaction"]
> [Tutorial de suplemento do Excel](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a>Veja também

* [Tutorial de suplemento do Excel](../tutorials/excel-tutorial-create-table.md)
* [Principais conceitos da API JavaScript do Excel](../excel/excel-add-ins-core-concepts.md)
* [Exemplos de código do suplemento do Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Referência da API JavaScript do Excel](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview?view=office-js)
