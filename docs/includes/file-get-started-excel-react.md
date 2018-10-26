# <a name="build-an-excel-add-in-using-react"></a>Criar um suplemento do Excel usando o React

Neste artigo, você percorrerá o processo de criação de um suplemento do Excel usando o React e a API JavaScript do Excel.

## <a name="prerequisites"></a>Pré-requisitos

- [Node.js](https://nodejs.org)

- Instale a última versão do [Yeoman](https://github.com/yeoman/yo) e o [gerador Yeoman para suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.
    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-web-app"></a>Criar o aplicativo Web

1. Use o gerador Yeoman para criar um projeto do suplemento do Excel. Execute o comando a seguir e responda às mensagens da seguinte forma:

    ```bash
    yo office
    ```

    - **Escolha um tipo de projeto:** `Office Add-in project using React framework`
    - **Como deseja nomear seu suplemento?:** `My Office Add-in`
    - **Qual aplicativo cliente do Office você gostaria de suportar?:** `Excel`

    ![Gerador do Yeoman](../images/yo-office-excel-react.png)
    
    Depois de concluir o assistente, o gerador criará o projeto e instalará os componentes de suporte do Node.

2. Navegue até a pasta raiz do projeto.

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a>Atualizar o código

1. No seu editor de código, abra o arquivo **src/styles.less**, adicione os seguintes estilos ao final do arquivo e salve o arquivo.

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
        font-family: Arial;
        padding-top: 25px;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
        font-family: Arial;
    }

    .padding {
        padding: 15px;
    }

    .padding-sm {
        padding: 4px;
    }

    .normal-button {
        width: 80px;
        padding: 2px;
    }
    ```

2. O modelo de projeto criado pelo gerador de Yeoman de Suplementos do Office inclui um componente React que não é necessário para esse início rápido. Exclua o arquivo **src/components/HeroList.tsx**.

3. Abra o arquivo **src/components/Header.tsx**, substitua todo o conteúdo com o código a seguir e salve o arquivo.

    ```typescript
    import * as React from 'react';

    export interface HeaderProps {
        title: string;
    }

    export class Header extends React.Component<HeaderProps, any> {
        constructor(props, context) {
            super(props, context);
        }

        render() {
            return (
                <div id='content-header'>
                    <div className='padding'>
                        <h1>{this.props.title}</h1>
                    </div>
                </div>
            );
        }
    }
    ```

4. Crie um novo componente React denominado **Content.tsx** na pasta **src/componentes** , adicione o código a seguir e salve o arquivo.

    ```typescript
    import * as React from 'react';
    import { Button, ButtonType } from 'office-ui-fabric-react';

    export interface ContentProps {
        message: string;
        buttonLabel: string;
        click: any;
    }

    export class Content extends React.Component<ContentProps, any> {
        constructor(props, context) {
            super(props, context);
        }

        render() {
            return (
                <div id='content-main'>
                    <div className='padding'>
                        <p>{this.props.message}</p>
                        <br />
                        <h3>Try it out</h3>
                        <br/>
                        <Button className='normal-button' buttonType={ButtonType.hero} onClick={this.props.click}>{this.props.buttonLabel}</Button>
                    </div>
                </div>
            );
        }
    }
    ```

5. Abra o arquivo **src/components/App.tsx**, substitua todo o conteúdo com o código a seguir e salve o arquivo.

    ```typescript
    import * as React from 'react';
    import { Header } from './Header';
    import { Content } from './Content';
    import Progress from './Progress';

    import * as OfficeHelpers from '@microsoft/office-js-helpers';

    export interface AppProps {
        title: string;
        isOfficeInitialized: boolean;
    }

    export interface AppState {
    }

    export default class App extends React.Component<AppProps, AppState> {
        constructor(props, context) {
            super(props, context);
        }

        setColor = async () => {
            try {
                await Excel.run(async context => {
                    const range = context.workbook.getSelectedRange();
                    range.load('address');
                    range.format.fill.color = 'green';
                    await context.sync();
                    console.log(`The range address was ${range.address}.`);
                });
            } catch (error) {
                OfficeHelpers.UI.notify(error);
                OfficeHelpers.Utilities.log(error);
            }
        }

        render() {
            const {
                title,
                isOfficeInitialized,
            } = this.props;

            if (!isOfficeInitialized) {
                return (
                    <Progress
                        title={title}
                        logo='assets/logo-filled.png'
                        message='Please sideload your addin to see app body.'
                    />
                );
            }

            return (
                <div className='ms-welcome'>
                    <Header title='Welcome' />
                    <Content message='Choose the button below to set the color of the selected range to green.' buttonLabel='Set color' click={this.setColor} />
                </div>
            );
        }
    }
    ```

## <a name="update-the-manifest"></a>Atualizar o manifesto

1. Abra o arquivo **manifest.xml** para definir as configurações e os recursos do suplemento. 

2. O elemento `ProviderName` tem um valor de espaço reservado. Substitua-o com seu nome.

3. O atributo `DefaultValue` do elemento `Description` tem um espaço reservado. Substitua-o por **Um suplemento do painel de tarefas para o Excel**.

4. Salve o arquivo.

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

## <a name="start-the-dev-server"></a>Iniciar o servidor de desenvolvimento

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a>Experimente

1. Siga as instruções para a plataforma que você usará para executar o suplemento e fazer o sideload do suplemento no Excel.

    - Windows: [Fazer o sideload de suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online: [Fazer o sideload dos suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad e Mac: [Fazer o sideload dos suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

2. No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-2b.png)

3. Selecione qualquer intervalo de células na planilha.

4. No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.

    ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>Próximas etapas

Parabéns, você criou com sucesso um suplemento do Excel usando o React! Em seguida, aprenda mais sobre os recursos de um suplemento do Excel e crie um suplemento mais complexo seguindo o tutorial do suplemento do Excel.

> [!div class="nextstepaction"]
> [Tutorial de suplementos do Excel](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a>Confira também

* [Tutorial de suplemento do Excel](../tutorials/excel-tutorial-create-table.md)
* [Conceitos fundamentais de programação com a API JavaScript do Excel](../excel/excel-add-ins-core-concepts.md)
* [Exemplos de código do suplemento do Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Referência da API JavaScript do Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js)
