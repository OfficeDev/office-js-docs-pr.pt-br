---
title: Usar o Office UI Fabric React em Suplementos do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 8ae8bac8c8043b51188d765dd7170922dcc1c84e
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="use-office-ui-fabric-react-in-office-add-ins"></a>Usar o Office UI Fabric React em Suplementos do Office

O Office UI Fabric ? uma estrutura de front-end JavaScript destinada ? cria??o de experi?ncias de usu?rio para Office e Office 365. Se voc? criar o suplemento usando o React, considere o uso do Fabric React para criar a experi?ncia do usu?rio. O Fabric fornece diversos componentes da experi?ncia de usu?rio baseados no React, como bot?es e caixas de sele??o, que voc? pode usar no suplemento.

Para come?ar a usar componentes do Fabric React no suplemento, execute as etapas a seguir.

> [!NOTE]
> Se voc? seguir as etapas nesta se??o, o Fabric Core tamb?m estar? dispon?vel no suplemento.

## <a name="step-1---create-your-project-with-the-yeoman-generator-for-office"></a>Etapa 1: criar o projeto com o gerador Yeoman para o Office

Para criar um suplemento que usa o Fabric React, recomendamos que voc? use o gerador Yeoman para o Office. O gerador Yeoman para o Office fornece o scaffolding de projeto e o gerenciamento de cria??o necess?rios para desenvolver um suplemento do Office.

Para criar o projeto, execute as seguintes etapas usando o **Windows PowerShell** (n?o o prompt de comando):

1. Instale os pr?-requisitos.
2. Execute o `yo office` para criar os arquivos de projeto para o suplemento.
3. Quando solicitado a selecionar um aplicativo cliente do Office, escolha o **Word**.
4. Voc? precisa estar no diret?rio com os arquivos de projeto e executar o `npm start`. Uma janela do navegador que mostra um controle girat?rio abrir? automaticamente.
5. [Execute sideload do manifesto](..\testing\test-debug-office-add-ins.md) para exibir a interface do usu?rio completa do suplemento.

## <a name="step-2---add-a-fabric-react-component"></a>Etapa 2: adicionar um componente do Fabric React

Em seguida, adicione componentes do Fabric React ao suplemento. Crie um novo componente do React, chamado `ButtonPrimaryExample`, que consiste em um Label e PrimaryButton do Fabric React. Para criar `ButtonPrimaryExample`:

1. Abra a pasta do projeto criada pelo gerador Yeoman e acesse **src\components**.
2. Crie **button.tsx**.
3. Em **button.tsx**, digite o c?digo a seguir para criar o componente `ButtonPrimaryExample`.

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

Esse c?digo faz o seguinte:

- Faz refer?ncia ? biblioteca React usando `import * as React from 'react';`.
- Faz refer?ncia aos componentes do Fabric (PrimaryButton, IButtonProps, Label) que s?o usados para criar `ButtonPrimaryExample`.
- Declara e torna p?blico o novo componente `ButtonPrimaryExample` usando `export class ButtonPrimaryExample extends React.Component`.
- Declara a fun??o `insertText` para manipular o evento `onClick`.
- Define a interface do usu?rio do componente do React na fun??o `render`. Renderiza e define a estrutura do componente. No `render`, voc? conecta o evento `onClick` usando `this.insertText`.

## <a name="step-3---add-the-react-component-to-your-add-in"></a>Etapa 3: adicionar o componente do React ao suplemento

Adicione `ButtonPrimaryExample` ao suplemento abrindo **src\components\app.tsx** e fazendo o seguinte:

- Adicione a seguinte instru??o de importa??o para fazer refer?ncia a `ButtonPrimaryExample` de **button.tsx** criado na etapa 2 (nenhuma extens?o de arquivo ? necess?ria).

  ```typescript
  import {ButtonPrimaryExample} from './button';
  ```

- Substitua a fun??o padr?o `render()` pelo c?digo a seguir, que usa `<ButtonPrimaryExample />`.

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

Salve suas altera??es. Todas as inst?ncias abertas do navegador, inclusive o suplemento, s?o atualizadas automaticamente e mostram o componente do React `ButtonPrimaryExample`. Observe que o texto padr?o e o bot?o s?o substitu?dos pelo texto e o bot?o principal definidos em `ButtonPrimaryExample`.

## <a name="recommended-components"></a>Componentes recomendados

Veja a seguir uma lista de componentes da experi?ncia de usu?rio do Fabric React recomendados para uso em suplementos:

- [Navega??o estrutural](breadcrumb.md)
- [Bot?o](button.md)
- [Caixa de sele??o](checkbox.md)
- [ChoiceGroup](choicegroup.md)
- [Lista suspensa](dropdown.md)
- [R?tulo](label.md)
- [Lista](list.md)
- [Tabela din?mica](pivot.md)
- [TextField](textfield.md)
- [Altern?ncia](toggle.md)

> [!NOTE]
> Adicionaremos outros componentes ao longo do tempo.

## <a name="see-also"></a>Veja tamb?m

- [Office UI Fabric React](https://dev.office.com/fabric#/)
- [Introdu??o ao exemplo de c?digo do Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [Padr?es de design da experi?ncia de usu?rio (usa o Fabric 2.6.1)](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Amostra de Fabric UI do suplemento do Office (usa o Fabric 1.0)](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [Gerador Yeoman para o Office](https://github.com/OfficeDev/generator-office)
