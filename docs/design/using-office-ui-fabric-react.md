---
title: Usar o Office UI Fabric React em Suplementos do Office
description: ''
ms.date: 12/04/2017
---
# <a name="use-office-ui-fabric-react-in-office-add-ins"></a>Usar o Office UI Fabric React em Suplementos do Office

O Office UI Fabric é uma estrutura de front-end JavaScript destinada à criação de experiências de usuário para Office e Office 365. Se você criar o suplemento usando o React, considere o uso do Fabric React para criar a experiência do usuário. O Fabric fornece diversos componentes da experiência de usuário baseados no React, como botões e caixas de seleção, que você pode usar no suplemento.

Para começar a usar componentes do Fabric React no suplemento, execute as etapas a seguir.

> [!NOTE]
> Se você seguir as etapas nesta seção, o Fabric Core também estará disponível no suplemento.

## <a name="step-1---create-your-project-with-the-yeoman-generator-for-office"></a>Etapa 1: criar o projeto com o gerador Yeoman para o Office

Para criar um suplemento que usa o Fabric React, recomendamos que você use o gerador Yeoman para o Office. O gerador Yeoman para o Office fornece o scaffolding de projeto e o gerenciamento de criação necessários para desenvolver um suplemento do Office.

Para criar o projeto, execute as seguintes etapas usando o **Windows PowerShell** (não o prompt de comando):

1. Instale os pré-requisitos.
2. Execute o `yo office` para criar os arquivos de projeto para o suplemento.
3. Quando solicitado a selecionar um aplicativo cliente do Office, escolha o **Word**.
4. Você precisa estar no diretório com os arquivos de projeto e executar o `npm start`. Uma janela do navegador que mostra um controle giratório abrirá automaticamente.
5. [Execute sideload do manifesto](..\testing\test-debug-office-add-ins.md) para exibir a interface do usuário completa do suplemento.

## <a name="step-2---add-a-fabric-react-component"></a>Etapa 2: adicionar um componente do Fabric React

Em seguida, adicione componentes do Fabric React ao suplemento. Crie um novo componente do React, chamado `ButtonPrimaryExample`, que consiste em um Label e PrimaryButton do Fabric React. Para criar `ButtonPrimaryExample`:

1. Abra a pasta do projeto criada pelo gerador Yeoman e acesse **src\components**.
2. Crie **button.tsx**.
3. Em **button.tsx**, digite o código a seguir para criar o componente `ButtonPrimaryExample`.

```typescript
import * as React from 'react';
import { PrimaryButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';

export class ButtonPrimaryExample extends React.Component<IButtonProps, {}> {
  public constructor() {
    super();
  }

   insertText = async () => {
        // In the click event, write text to the document.
        await Word.run(async (context) => {
            let body = context.document.body;
            body.insertParagraph('Hello Office UI Fabric React!', Word.InsertLocation.end);
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
- Faz referência aos componentes do Fabric (PrimaryButton, IButtonProps, Label) que são usados para criar `ButtonPrimaryExample`.
- Declara e torna público o novo componente `ButtonPrimaryExample` usando `export class ButtonPrimaryExample extends React.Component`.
- Declara a função `insertText` para manipular o evento `onClick`.
- Define a interface do usuário do componente do React na função `render`. Renderiza e define a estrutura do componente. No `render`, você conecta o evento `onClick` usando `this.insertText`.

## <a name="step-3---add-the-react-component-to-your-add-in"></a>Etapa 3: adicionar o componente do React ao suplemento

Adicione `ButtonPrimaryExample` ao suplemento abrindo **src\components\app.tsx** e fazendo o seguinte:

- Adicione a seguinte instrução de importação para fazer referência a `ButtonPrimaryExample` de **button.tsx** criado na etapa 2 (nenhuma extensão de arquivo é necessária).

  ```typescript
  import {ButtonPrimaryExample} from './button';
  ```

- Substitua a função padrão `render()` pelo código a seguir, que usa `<ButtonPrimaryExample />`.

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

Salve suas alterações. Todas as instâncias abertas do navegador, inclusive o suplemento, são atualizadas automaticamente e mostram o componente do React `ButtonPrimaryExample`. Observe que o texto padrão e o botão são substituídos pelo texto e o botão principal definidos em `ButtonPrimaryExample`.

## <a name="recommended-components"></a>Componentes recomendados

Veja a seguir uma lista de componentes da experiência de usuário do Fabric React recomendados para uso em suplementos:

- [Navegação estrutural](breadcrumb.md)
- [Botão](button.md)
- [Caixa de seleção](checkbox.md)
- [ChoiceGroup](choicegroup.md)
- [Lista suspensa](dropdown.md)
- [Rótulo](label.md)
- [Lista](list.md)
- [Tabela dinâmica](pivot.md)
- [Campo de texto](textfield.md)
- [Alternância](toggle.md)

> [!NOTE]
> Adicionaremos outros componentes ao longo do tempo.

## <a name="see-also"></a>Veja também

- [Office UI Fabric React](https://dev.office.com/fabric#/)
- [Introdução ao exemplo de código do Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [Padrões de design da experiência de usuário (usa o Fabric 2.6.1)](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Amostra de Fabric UI do suplemento do Office (usa o Fabric 1.0)](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [Usar o Fabric 2.6.1 em um Suplemento do Office](ui-elements/using-office-ui-fabric.md)
- [Gerador Yeoman para o Office](https://github.com/OfficeDev/generator-office)
