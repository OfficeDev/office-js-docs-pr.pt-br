---
title: Interface do usuário Fluent React em Suplementos do Office
description: Saiba como usar Fluent interface do usuário React em Office de complementos.
ms.date: 01/13/2022
ms.localizationpriority: medium
ms.openlocfilehash: 453befe44dbcec6527930fcd73c5cb2cb243d965
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743130"
---
# <a name="use-fluent-ui-react-in-office-add-ins"></a>Usar Fluent interface do usuário React em Office de usuário

Fluent interface do usuário React é a estrutura de front-end javaScript de código aberto oficial projetada para criar experiências que se encaixem perfeitamente em uma ampla variedade de produtos Microsoft, incluindo Office. Ele fornece componentes robustos, atualizados e acessíveis baseados no React que são altamente personalizáveis usando o CSS-in-JS.

> [!NOTE]
> Este artigo descreve o uso de Fluent de interface do usuário React no contexto de Office Desem. Mas também é usado em uma ampla variedade de Microsoft 365 aplicativos e extensões. Para obter mais informações, [consulte Fluent interface do usuário React](https://developer.microsoft.com/fluentui#/get-started/web#fluent-ui-react) e o aplicativo de Fluent [da interface do usuário.](https://github.com/microsoft/fluentui)

Este artigo descreve como criar um add-in criado com o React e usa Fluent de interface do usuário React componentes.

## <a name="create-an-add-in-project"></a>Criar um projeto de suplemento

Você usará o gerador Yeoman para Suplementos do Office para criar um projeto de suplemento que use o React.

### <a name="install-the-prerequisites"></a>Instalar os pré-requisitos

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-project"></a>Criar o projeto

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Escolha o tipo de projeto:** `Office Add-in Task Pane project using React framework`
- **Escolha o tipo de script:** `TypeScript`
- **Qual será o nome do suplemento?** `My Office Add-in`
- **Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?** `Word`

![Captura de tela apresentando os avisos e respostas do gerador Yeoman em uma interface de linha de comando.](../images/yo-office-word-react.png)

Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="try-it-out"></a>Experimente

1. Navegue até a pasta raiz do projeto.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. Conclua as etapas a seguir para iniciar o servidor da web local e fazer o sideload do seu suplemento.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    > [!TIP]
    > Se você estiver testando o seu suplemento no Mac, execute o seguinte comando antes de continuar. O servidor Web local é iniciado quando este comando é executado.
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - Para testar seu suplemento no Word, execute o seguinte comando no diretório raiz do seu projeto. Isso inicia o servidor Web local e abre o Word com o seu complemento carregado.

        ```command&nbsp;line
        npm start
        ```

    - Para testar seu suplemento no Word em um navegador, execute o seguinte comando no diretório raiz do seu projeto. O servidor Web local é iniciado quando este comando é executado. Substitua "{url}" pelo URL de um documento do Word no seu OneDrive ou uma biblioteca do SharePoint para a qual você tenha permissões.

        [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

3. Para abrir o painel de tarefas do complemento, na guia **Página** Início, escolha o **botão Mostrar Painel de Tarefas** . Observe o texto padrão e o botão **Executar** na parte inferior do painel de tarefas. No restante deste passo a passo, você redefinirá esse texto e um botão criando um componente React que usa componentes de interface do usuário Fluent interface do usuário React.

    ![Captura de tela mostrando o aplicativo Word com o botão mostrar faixa de opções do Painel de Tarefas realçada e o botão Executar e imediatamente o texto anterior realçado no painel de tarefas.](../images/word-task-pane-yo-default.png)

## <a name="create-a-react-component-that-uses-fluent-ui-react"></a>Criar um React que usa Fluent interface do usuário React

Neste ponto, você criou um suplemento muito básico do painel de tarefas usando o React. Em seguida, siga as etapas abaixo para criar um novo componente React (`ButtonPrimaryExample`) dentro do projeto de suplemento. O componente usa os componentes `Label` e `PrimaryButton` de Fluent interface do usuário React.

1. Abra a pasta do projeto criada pelo gerador Yeoman e acesse **src\taskpane\components**.
2. Nesta pasta, crie um novo arquivo chamado **Button.tsx**.
3. Em **Button.tsx**, adicione o código a seguir para definir o componente `ButtonPrimaryExample`.

```typescript
import * as React from 'react';
import { PrimaryButton, IButtonProps } from '@fluentui/react/lib/components/Button';
import { Label } from '@fluentui/react/lib/components/Label';

export class ButtonPrimaryExample extends React.Component<IButtonProps, {}> {
  public constructor(props) {
    super(props);
  }

  insertText = async () => {
    // In the click event, write text to the document.
    await Word.run(async (context) => {
      let body = context.document.body;
      body.insertParagraph('Hello Fluent UI React!', Word.InsertLocation.end);
      await context.sync();
    });
  }

  public render() {
    let { disabled } = this.props;
    return (
      <div className='ms-BasicButtonsExample'>
        <Label>Click the button to insert text.</Label>
        <PrimaryButton
          data-automation-id='test'
          disabled={ disabled }
          text='Insert text...'
          onClick={ this.insertText } />
      </div>
    );
  }
}
```

Esse código faz o seguinte:

- Faz referência à biblioteca React usando `import * as React from 'react';`.
- Faz referência aos Fluent de interface do usuário React (`PrimaryButton`, , ), `Label`que são usados para criar `ButtonPrimaryExample``IButtonProps`.
- Declara o novo `ButtonPrimaryExample` componente usando `export class ButtonPrimaryExample extends React.Component`.
- Declara a `insertText` função que manipulará o evento do `onClick` botão.
- Define a interface do usuário do componente do React na função `render`. A marcação HTML usa os `Label` `PrimaryButton` `onClick` componentes e do Fluent interface do usuário React e especifica que, quando o evento for ativos, `insertText` a função será executado.

## <a name="add-the-react-component-to-your-add-in"></a>Adicionar o componente do React ao suplemento

Adicione o `ButtonPrimaryExample` componente ao seu complemento abrindo **src\components\App.tsx** e concluindo as etapas a seguir.

1. Adicione a seguinte declaração de importação para a referência `ButtonPrimaryExample` do **Button.tsx**.

    ```typescript
    import {ButtonPrimaryExample} from './Button';
    ```

2. Remova a instrução de importação a seguir.

    ```typescript
    import Progress from './Progress';
    ```

3. Substitua a função padrão `render()` pelo código a seguir, que usa `ButtonPrimaryExample`.

    ```typescript
    render() {
      return (
        <div className="ms-welcome">
        <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome" />
        <HeroList message="Discover what this add-in can do for you today!" items={this.state.listItems} >
          <ButtonPrimaryExample />
        </HeroList>
        </div>
      );
    }
    ```

4. Salve as alterações feitas em **App.tsx**.

## <a name="see-the-result"></a>Ver o resultado

No Word, o painel de tarefas do suplemento será atualizado automaticamente quando você salvar as alterações em **App.tsx**. O texto padrão e o botão na parte inferior do painel de tarefas agora mostram a IU definida pelo `ButtonPrimaryExample` componente. Feche o botão **Insert text...** para inserir o texto no documento.

![Captura de tela mostrando o aplicativo Word com o "Inserir texto..." botão e texto imediatamente anterior realçado.](../images/word-task-pane-with-react-component.png)

Parabéns, você criou com êxito um complemento do painel de tarefas usando o React e Fluent interface do usuário React!

## <a name="see-also"></a>Confira também

- [Word Add-in GettingStartedFabricReact](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [Núcleo da Malha em Suplementos do Office](fabric-core.md)
- [Padrões de design da experiência do usuário para suplementos do Office](ux-design-pattern-templates.md)
