---
title: Usar o Office UI Fabric React em Suplementos do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 4baeea20457892bcc7b94b381f5c0a577274408a
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944272"
---
# <a name="use-office-ui-fabric-react-in-office-add-ins"></a><span data-ttu-id="2303a-102">Usar o Office UI Fabric React em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="2303a-102">Use Office UI Fabric React in Office Add-ins</span></span>

<span data-ttu-id="2303a-p101">O Office UI Fabric é uma estrutura de front-end JavaScript destinada à criação de experiências de usuário para Office e Office 365. Se você criar o suplemento usando o React, considere o uso do Fabric React para criar a experiência do usuário. O Fabric fornece diversos componentes da experiência de usuário baseados no React, como botões e caixas de seleção, que você pode usar no suplemento.</span><span class="sxs-lookup"><span data-stu-id="2303a-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. If you build your add-in using React, consider using Fabric React to create your user experience. Fabric provides several React-based UX components, like buttons or checkboxes, that you can use in your add-in.</span></span>

<span data-ttu-id="2303a-106">Para começar a usar componentes do Fabric React no suplemento, execute as etapas a seguir.</span><span class="sxs-lookup"><span data-stu-id="2303a-106">To get started using Fabric React's components in your add-in, perform the following steps.</span></span>

> [!NOTE]
> <span data-ttu-id="2303a-107">Se você seguir as etapas nesta seção, o Fabric Core também estará disponível no suplemento.</span><span class="sxs-lookup"><span data-stu-id="2303a-107">If you follow the steps in this article, Fabric Core is also available in your add-in.</span></span>

## <a name="step-1---create-your-project-with-the-yeoman-generator-for-office"></a><span data-ttu-id="2303a-108">Etapa 1: criar o projeto com o gerador Yeoman para o Office</span><span class="sxs-lookup"><span data-stu-id="2303a-108">Step 1 - Create your project with the Yeoman generator for Office</span></span>

<span data-ttu-id="2303a-p102">Para criar um suplemento que usa o Fabric React, recomendamos que você use o gerador Yeoman para Office. O gerador Yeoman para Office fornece o scaffolding de projeto e o gerenciamento de criação necessários para desenvolver um suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="2303a-p102">To create an add-in that uses Fabric React, we recommend that you use the Yeoman generator for Office. The Yeoman generator for Office provides the project scaffolding and build management needed to develop an Office add-in.</span></span>

<span data-ttu-id="2303a-111">Para criar o projeto, execute as seguintes etapas usando o **Windows PowerShell** (não o prompt de comando):</span><span class="sxs-lookup"><span data-stu-id="2303a-111">To create your project, perform the following steps using **Windows PowerShell** (not the command prompt):</span></span>

1. <span data-ttu-id="2303a-112">Instale os pré-requisitos.</span><span class="sxs-lookup"><span data-stu-id="2303a-112">Install the prerequisites.</span></span>
2. <span data-ttu-id="2303a-113">Execute o `yo office` para criar os arquivos de projeto para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="2303a-113">Run `yo office` to create the project files for your add-in.</span></span>
3. <span data-ttu-id="2303a-114">Quando solicitado a selecionar um aplicativo cliente do Office, escolha o **Word**.</span><span class="sxs-lookup"><span data-stu-id="2303a-114">When prompted to select an Office client application, choose **Word**.</span></span>
4. <span data-ttu-id="2303a-p103">Você precisa estar no diretório com os arquivos de projeto e executar o `npm start`. Uma janela do navegador que mostra um controle giratório abrirá automaticamente.</span><span class="sxs-lookup"><span data-stu-id="2303a-p103">Ensure you are in the directory with the project files, and then run `npm start`. A browser window showing a spinner opens automatically.</span></span>
5. <span data-ttu-id="2303a-117">[Execute sideload do manifesto](..\testing\test-debug-office-add-ins.md) para exibir a interface do usuário completa do suplemento.</span><span class="sxs-lookup"><span data-stu-id="2303a-117">[Sideload your manifest](..\testing\test-debug-office-add-ins.md) to view the full UI of the add-in.</span></span>

## <a name="step-2---add-a-fabric-react-component"></a><span data-ttu-id="2303a-118">Etapa 2: adicionar um componente do Fabric React</span><span class="sxs-lookup"><span data-stu-id="2303a-118">Step 2 - Add a Fabric React component</span></span>

<span data-ttu-id="2303a-p104">Em seguida, adicione componentes do Fabric React ao suplemento. Crie um novo componente do React, chamado `ButtonPrimaryExample`, que consiste em um Label e PrimaryButton do Fabric React. Para criar `ButtonPrimaryExample`:</span><span class="sxs-lookup"><span data-stu-id="2303a-p104">Next, add Fabric React components to your add-in. Create a new React component, called `ButtonPrimaryExample`, that consists of a Label and PrimaryButton from Fabric React. To create `ButtonPrimaryExample`:</span></span>

1. <span data-ttu-id="2303a-122">Abra a pasta do projeto criada pelo gerador Yeoman e acesse **src\components**.</span><span class="sxs-lookup"><span data-stu-id="2303a-122">Open the project folder created by the Yeoman generator, and go to **src\components**.</span></span>
2. <span data-ttu-id="2303a-123">Crie **button.tsx**.</span><span class="sxs-lookup"><span data-stu-id="2303a-123">Create **button.tsx**.</span></span>
3. <span data-ttu-id="2303a-124">Em **button.tsx**, digite o código a seguir para criar o componente `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="2303a-124">In **button.tsx**, enter the following code to create the `ButtonPrimaryExample` component.</span></span>

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

<span data-ttu-id="2303a-125">Esse código faz o seguinte:</span><span class="sxs-lookup"><span data-stu-id="2303a-125">This code does the following:</span></span>

- <span data-ttu-id="2303a-126">Faz referência à biblioteca React usando `import * as React from 'react';`.</span><span class="sxs-lookup"><span data-stu-id="2303a-126">References the React library using `import * as React from 'react';`.</span></span>
- <span data-ttu-id="2303a-127">Faz referência aos componentes do Fabric (PrimaryButton, IButtonProps, Label) que são usados para criar `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="2303a-127">References the Fabric components (PrimaryButton, IButtonProps, Label) that are used to create `ButtonPrimaryExample`.</span></span>
- <span data-ttu-id="2303a-128">Declara e torna público o novo componente `ButtonPrimaryExample` usando `export class ButtonPrimaryExample extends React.Component`.</span><span class="sxs-lookup"><span data-stu-id="2303a-128">Declares and make public the new `ButtonPrimaryExample` component using `export class ButtonPrimaryExample extends React.Component`.</span></span>
- <span data-ttu-id="2303a-129">Declara a função `insertText` para manipular o evento `onClick`.</span><span class="sxs-lookup"><span data-stu-id="2303a-129">Declares the `insertText` function to handle the `onClick` event.</span></span>
- <span data-ttu-id="2303a-p105">Define a interface do usuário do componente do React na função `render`. Renderiza e define a estrutura do componente. No `render`, você conecta o evento `onClick` usando `this.insertText`.</span><span class="sxs-lookup"><span data-stu-id="2303a-p105">Defines the UI of the React component in the `render` function. Render defines the structure of the component. Within `render`, you wire up the `onClick` event using `this.insertText`.</span></span>

## <a name="step-3---add-the-react-component-to-your-add-in"></a><span data-ttu-id="2303a-133">Etapa 3: adicionar o componente do React ao suplemento</span><span class="sxs-lookup"><span data-stu-id="2303a-133">Step 3 - Add the React component to your add-in</span></span>

<span data-ttu-id="2303a-134">Adicione `ButtonPrimaryExample` ao suplemento abrindo **src\components\app.tsx** e fazendo o seguinte:</span><span class="sxs-lookup"><span data-stu-id="2303a-134">Add `ButtonPrimaryExample` to your add-in by opening **src\components\app.tsx** and doing the following:</span></span>

- <span data-ttu-id="2303a-135">Adicione a seguinte instrução de importação para fazer referência a `ButtonPrimaryExample` de **button.tsx** criado na etapa 2 (nenhuma extensão de arquivo é necessária).</span><span class="sxs-lookup"><span data-stu-id="2303a-135">Add the following import statement to reference `ButtonPrimaryExample` from **button.tsx** created in step 2 (no file extension is needed).</span></span>

  ```typescript
  import {ButtonPrimaryExample} from './button';
  ```

- <span data-ttu-id="2303a-136">Substitua a função padrão `render()` pelo código a seguir, que usa `<ButtonPrimaryExample />`.</span><span class="sxs-lookup"><span data-stu-id="2303a-136">Replace the default `render()` function with the following code that uses `<ButtonPrimaryExample />`.</span></span>

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

<span data-ttu-id="2303a-p106">Salve suas alterações. Todas as instâncias abertas do navegador, inclusive o suplemento, são atualizadas automaticamente e mostram o componente do React `ButtonPrimaryExample`. Observe que o texto padrão e o botão são substituídos pelo texto e o botão principal definidos em `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="2303a-p106">Save your changes. All open browser instances, including the add-in, update automatically and show the `ButtonPrimaryExample` React component. Notice that the default text and button is replaced with the text and primary button defined in `ButtonPrimaryExample`.</span></span>



## <a name="see-also"></a><span data-ttu-id="2303a-140">Confira também</span><span class="sxs-lookup"><span data-stu-id="2303a-140">See also</span></span>

- [<span data-ttu-id="2303a-141">Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="2303a-141">Office UI Fabric React</span></span>](https://developer.microsoft.com/fabric)
- [<span data-ttu-id="2303a-142">Introdução ao exemplo de código do Fabric React</span><span class="sxs-lookup"><span data-stu-id="2303a-142">Getting started with Fabric React code sample</span></span>](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [<span data-ttu-id="2303a-143">Padrões de design da experiência de usuário (usa o Fabric 2.6.1)</span><span class="sxs-lookup"><span data-stu-id="2303a-143">UX design patterns (uses Fabric 2.6.1)</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="2303a-144">Exemplo de suplemento do Office com Fabric UI (usa o Fabric 1.0)</span><span class="sxs-lookup"><span data-stu-id="2303a-144">Office Add-in Fabric UI sample (uses Fabric 1.0)</span></span>](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [<span data-ttu-id="2303a-145">Gerador Yeoman para Office</span><span class="sxs-lookup"><span data-stu-id="2303a-145">Yeoman generator for Office</span></span>](https://github.com/OfficeDev/generator-office)
