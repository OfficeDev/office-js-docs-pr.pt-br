---
title: Usar o Office UI Fabric React em Suplementos do Office
description: Aprenda a usar o Office UI Fabric React em suplementos do Office.
ms.date: 07/11/2019
localization_priority: Priority
ms.openlocfilehash: c738521b82d0cb8f234fd28dc8bb24740962b817
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/13/2019
ms.locfileid: "36302593"
---
# <a name="use-office-ui-fabric-react-in-office-add-ins"></a><span data-ttu-id="e2339-103">Usar o Office UI Fabric React em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="e2339-103">Use Office UI Fabric React in Office Add-ins</span></span>

<span data-ttu-id="e2339-p101">O Office UI Fabric é uma estrutura de front-end JavaScript destinada à criação de experiências de usuário para Office e Office 365. Se você criar o suplemento usando o React, considere o uso do Fabric React para criar a experiência do usuário. O Fabric fornece diversos componentes da experiência de usuário baseados no React, como botões e caixas de seleção, que você pode usar no suplemento.</span><span class="sxs-lookup"><span data-stu-id="e2339-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. If you build your add-in using React, consider using Fabric React to create your user experience. Fabric provides several React-based UX components, like buttons or checkboxes, that you can use in your add-in.</span></span>

<span data-ttu-id="e2339-107">Este artigo descreve como criar um suplemento usando o React e componentes do Fabric React.</span><span class="sxs-lookup"><span data-stu-id="e2339-107">This article describes how to create an add-in that's built with React and uses Fabric React components.</span></span> 

> [!NOTE]
> <span data-ttu-id="e2339-108">[O Fabric Core](office-ui-fabric.md#use-fabric-core-icons-fonts-colors) acompanha o Fabric React, o que significa que o seu suplemento também terá acesso ao Fabric Core após a conclusão das etapas deste artigo.</span><span class="sxs-lookup"><span data-stu-id="e2339-108">[Fabric Core](office-ui-fabric.md#use-fabric-core-icons-fonts-colors) is included with Fabric React, which means your add-in will also have access to Fabric Core after you've completed the steps in this article.</span></span>

## <a name="create-an-add-in-project"></a><span data-ttu-id="e2339-109">Criar um projeto de suplemento</span><span class="sxs-lookup"><span data-stu-id="e2339-109">Create an Outlook add-in project</span></span>

<span data-ttu-id="e2339-110">Você usará o gerador Yeoman para Suplementos do Office para criar um projeto de suplemento que use o React.</span><span class="sxs-lookup"><span data-stu-id="e2339-110">You'll use the Yeoman generator for Office Add-ins to create an add-in project that uses React.</span></span>

### <a name="install-the-prerequisites"></a><span data-ttu-id="e2339-111">Instalar os pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="e2339-111">Install the prerequisites.</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-project"></a><span data-ttu-id="e2339-112">Criar o projeto</span><span class="sxs-lookup"><span data-stu-id="e2339-112">Create the project</span></span>

[!include[note about Yeoman generator bug](../includes/note-yeoman-generator-bug-201908.md)]

<span data-ttu-id="e2339-113">Use o gerador Yeoman para criar um projeto de suplemento do Word.</span><span class="sxs-lookup"><span data-stu-id="e2339-113">Use the Yeoman generator to create a Word add-in project.</span></span> <span data-ttu-id="e2339-114">Execute o comando a seguir e responda aos prompts da seguinte forma:</span><span class="sxs-lookup"><span data-stu-id="e2339-114">Run the following command and then answer the prompts as follows:</span></span>

```command&nbsp;line
yo office
```

- <span data-ttu-id="e2339-115">**Escolha o tipo de projeto:** `Office Add-in Task Pane project using React framework`</span><span class="sxs-lookup"><span data-stu-id="e2339-115">**Choose a project type:** `Office Add-in Task Pane project using React framework`</span></span>
- <span data-ttu-id="e2339-116">**Escolha o tipo de script:** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="e2339-116">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="e2339-117">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="e2339-117">**What do you want to name your add-in?**</span></span> `my-office-add-in`
- <span data-ttu-id="e2339-118">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?**</span><span class="sxs-lookup"><span data-stu-id="e2339-118">**Which Office client application would you like to support?**</span></span> `Word`

<span data-ttu-id="e2339-119">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="e2339-119">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

### <a name="try-it-out"></a><span data-ttu-id="e2339-120">Experimente</span><span class="sxs-lookup"><span data-stu-id="e2339-120">Try it out</span></span>

1. <span data-ttu-id="e2339-121">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="e2339-121">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "my-office-add-in"
    ```

2. <span data-ttu-id="e2339-122">Conclua as etapas a seguir para iniciar o servidor Web local e fazer o sideload do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="e2339-122">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="e2339-123">Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="e2339-123">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="e2339-124">Se você for solicitado a instalar um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.</span><span class="sxs-lookup"><span data-stu-id="e2339-124">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    > [!TIP]
    > <span data-ttu-id="e2339-125">Se você estiver testando o seu suplemento no Mac, execute o seguinte comando antes de continuar.</span><span class="sxs-lookup"><span data-stu-id="e2339-125">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="e2339-126">O servidor Web local é iniciado quando este comando é executado.</span><span class="sxs-lookup"><span data-stu-id="e2339-126">When you run this command, the local web server will start.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="e2339-127">Para testar seu suplemento no Word, execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="e2339-127">To test your add-in in Word, run the following command in the root directory of your project.</span></span> <span data-ttu-id="e2339-128">Isso inicia o servidor Web local (caso ainda não esteja em execução) e abre o Word com o suplemento carregado.</span><span class="sxs-lookup"><span data-stu-id="e2339-128">This starts the local web server (if it's not already running) and opens Word with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="e2339-129">Para testar seu suplemento no Word em um navegador, execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="e2339-129">To test your add-in in Word on a browser, run the following command in the root directory of your project.</span></span> <span data-ttu-id="e2339-130">Quando você executar este comando, o servidor Web local será iniciado (se ainda não estiver em execução).</span><span class="sxs-lookup"><span data-stu-id="e2339-130">When you run this command, the local web server will start.</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="e2339-131">Para usar o seu suplemento, abra um novo documento no Word na Web e em seguida realize o sideload no suplemento de acordo com as instruções em [Realizar Sideload nos Suplementos do Office no Office na Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="e2339-131">To use your add-in, open a new document in Word Online and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

3. <span data-ttu-id="e2339-132">No Word, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="e2339-132">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="e2339-133">Observe o texto padrão e o botão **Executar** na parte inferior do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="e2339-133">Notice the default text and the **Run** button at the bottom of the task pane.</span></span> <span data-ttu-id="e2339-134">No restante deste passo a passo, você redefinirá esse texto e botão criando um componente Reagir que usa componentes UX do Fabric React.</span><span class="sxs-lookup"><span data-stu-id="e2339-134">In the remainder of this walkthrough, you'll redefine this text and button by creating a React component that uses UX components from Fabric React.</span></span>

    ![Captura de tela do aplicativo do Word com o botão Mostrar faixa de opções do painel de tarefas realçado e o botão Executar... e o texto anterior realçado no painel de tarefas](../images/word-task-pane-yo-default.png)


## <a name="create-a-react-component-that-uses-fabric-react"></a><span data-ttu-id="e2339-136">Criar um componente React que use o Fabric React</span><span class="sxs-lookup"><span data-stu-id="e2339-136">Create a React component that uses Fabric React</span></span>

<span data-ttu-id="e2339-137">Neste ponto, você criou um suplemento muito básico do painel de tarefas usando o React.</span><span class="sxs-lookup"><span data-stu-id="e2339-137">At this point, you've created a very basic task pane add-in that's built using React.</span></span> <span data-ttu-id="e2339-138">Em seguida, siga as etapas abaixo para criar um novo componente React (`ButtonPrimaryExample`) dentro do projeto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="e2339-138">Next, complete the following steps to create a new React component (`ButtonPrimaryExample`) within the add-in project.</span></span> <span data-ttu-id="e2339-139">O componente usa o `Label` e `PrimaryButton` os componentes do Fabric React.</span><span class="sxs-lookup"><span data-stu-id="e2339-139">The component uses the `Label` and `PrimaryButton` components from Fabric React.</span></span>

1. <span data-ttu-id="e2339-140">Abra a pasta do projeto criada pelo gerador Yeoman e acesse **src\taskpane\components**.</span><span class="sxs-lookup"><span data-stu-id="e2339-140">Open the project folder created by the Yeoman generator, and go to **src\components**.</span></span>
2. <span data-ttu-id="e2339-141">Nesta pasta, crie um novo arquivo chamado **Button.tsx**.</span><span class="sxs-lookup"><span data-stu-id="e2339-141">In that folder, create a new file named **Button.tsx**.</span></span>
3. <span data-ttu-id="e2339-142">Em **Button.tsx**, adicione o código a seguir para definir o componente `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="e2339-142">In **Button.tsx**, add the following code to define the `ButtonPrimaryExample` component.</span></span>

```typescript
import * as React from 'react';
import { PrimaryButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';

export class ButtonPrimaryExample extends React.Component<IButtonProps, {}> {
  public constructor(props) {
    super(props);
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

<span data-ttu-id="e2339-143">Esse código faz o seguinte:</span><span class="sxs-lookup"><span data-stu-id="e2339-143">This code does the following:</span></span>

- <span data-ttu-id="e2339-144">Faz referência à biblioteca React usando `import * as React from 'react';`.</span><span class="sxs-lookup"><span data-stu-id="e2339-144">References the React library using `import * as React from 'react';`.</span></span>
- <span data-ttu-id="e2339-145">Faz referência aos componentes do Fabric (`PrimaryButton`, `IButtonProps`, `Label`) que são usados para criar `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="e2339-145">References the Fabric components (PrimaryButton, IButtonProps, Label) that are used to create .</span></span>
- <span data-ttu-id="e2339-146">Declara o novo `ButtonPrimaryExample` componente usando `export class ButtonPrimaryExample extends React.Component`.</span><span class="sxs-lookup"><span data-stu-id="e2339-146">Declares and make public the new `ButtonPrimaryExample` component using `export class ButtonPrimaryExample extends React.Component`.</span></span>
- <span data-ttu-id="e2339-147">Declara a `insertText` função que manipulará o evento do `onClick` botão.</span><span class="sxs-lookup"><span data-stu-id="e2339-147">Declares the `insertText` function that will handle the button's `onClick` event.</span></span>
- <span data-ttu-id="e2339-148">Define a interface do usuário do componente do React na função `render`.</span><span class="sxs-lookup"><span data-stu-id="e2339-148">Defines the UI of the React component in the `render` function.</span></span> <span data-ttu-id="e2339-149">A marcação HTML usa os componentes `Label` e `PrimaryButton` da Fabric React e especifica que quando `onClick` o evento for acionado, a `insertText` função será executada.</span><span class="sxs-lookup"><span data-stu-id="e2339-149">The HTML markup uses the `Label` and `PrimaryButton` components from Fabric React and specifies that when the `onClick` event fires, the `insertText` function will run.</span></span>

## <a name="add-the-react-component-to-your-add-in"></a><span data-ttu-id="e2339-150">Adicionar o componente do React ao suplemento</span><span class="sxs-lookup"><span data-stu-id="e2339-150">Step 3 - Add the React component to your add-in</span></span>

<span data-ttu-id="e2339-151">Adicionar o `ButtonPrimaryExample` componente ao suplemento abrindo **src\components\App.tsx** e seguir as seguintes etapas:</span><span class="sxs-lookup"><span data-stu-id="e2339-151">Add the `ButtonPrimaryExample` component to your add-in by opening **src\components\App.tsx** and completing the following steps:</span></span>

1. <span data-ttu-id="e2339-152">Adicione a seguinte declaração de importação para a referência `ButtonPrimaryExample` do **Button.tsx**.</span><span class="sxs-lookup"><span data-stu-id="e2339-152">Add the following import statement to reference `ButtonPrimaryExample` from **Button.tsx**.</span></span>

    ```typescript
    import {ButtonPrimaryExample} from './Button';
    ```

2. <span data-ttu-id="e2339-153">Remova as duas instruções de importação a seguir.</span><span class="sxs-lookup"><span data-stu-id="e2339-153">Remove the following two import statements.</span></span>

    ```typescript
    import { Button, ButtonType } from 'office-ui-fabric-react';
    ...
    import Progress from './Progress';
    ```

3. <span data-ttu-id="e2339-154">Substitua a função padrão `render()` pelo código a seguir, que usa `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="e2339-154">Replace the default `render()` function with the following code that uses `ButtonPrimaryExample`.</span></span>

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

  4. <span data-ttu-id="e2339-155">Salve as alterações feitas em **App.tsx**.</span><span class="sxs-lookup"><span data-stu-id="e2339-155">Save the changes you've made to **App.tsx**.</span></span>

## <a name="see-the-result"></a><span data-ttu-id="e2339-156">Ver o resultado</span><span class="sxs-lookup"><span data-stu-id="e2339-156">See the result</span></span>

<span data-ttu-id="e2339-157">No Word, o painel de tarefas do suplemento será atualizado automaticamente quando você salvar as alterações em **App.tsx**.</span><span class="sxs-lookup"><span data-stu-id="e2339-157">In Word, the add-in task pane automatically updates when you save changes to **App.tsx**.</span></span> <span data-ttu-id="e2339-158">O texto padrão e o botão na parte inferior do painel de tarefas agora mostram a IU definida pelo `ButtonPrimaryExample` componente.</span><span class="sxs-lookup"><span data-stu-id="e2339-158">The default text and button at the bottom of the task pane now shows the UI that's defined by the `ButtonPrimaryExample` component.</span></span> <span data-ttu-id="e2339-159">Feche o botão **Insert text...** para inserir o texto no documento.</span><span class="sxs-lookup"><span data-stu-id="e2339-159">Choose the **Insert text...** button to insert text into the document.</span></span>

![Captura de tela do aplicativo Word com o botão Inserir texto... e o texto anterior realçado](../images/word-task-pane-with-react-component.png)

<span data-ttu-id="e2339-161">Parabéns, você criou com êxito um suplemento do painel de tarefas usando React e o Office UI Fabric React!</span><span class="sxs-lookup"><span data-stu-id="e2339-161">Congratulations, you've successfully created a task pane add-in using React and Office UI Fabric React!</span></span> 

## <a name="see-also"></a><span data-ttu-id="e2339-162">Confira também</span><span class="sxs-lookup"><span data-stu-id="e2339-162">See also</span></span>

- [<span data-ttu-id="e2339-163">Office UI Fabric em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="e2339-163">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md)
- [<span data-ttu-id="e2339-164">Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="e2339-164">Office UI Fabric React</span></span>](https://developer.microsoft.com/fabric)
- [<span data-ttu-id="e2339-165">Padrões de design da experiência do usuário para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="e2339-165">UX design patterns for Office Add-ins</span></span>](ux-design-pattern-templates.md)
- [<span data-ttu-id="e2339-166">Introdução ao exemplo de código do Fabric React</span><span class="sxs-lookup"><span data-stu-id="e2339-166">Getting started with Fabric React code sample</span></span>](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
