---
title: Interface do usuário do Fluent React em Office de complementos
description: Saiba como usar a interface do usuário do Fluent React em Office-ins.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: cb7f04c21a52a2e4a3f271abc56aa325dd2b02fd
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330134"
---
# <a name="use-fluent-ui-react-in-office-add-ins"></a><span data-ttu-id="96e70-103">Usar a interface do usuário do Fluent React em Office de complementos</span><span class="sxs-lookup"><span data-stu-id="96e70-103">Use Fluent UI React in Office Add-ins</span></span>

<span data-ttu-id="96e70-104">A interface do usuário React é a estrutura de front-end javaScript de código aberto oficial projetada para criar experiências que se encaixem perfeitamente em uma ampla variedade de produtos Microsoft, incluindo Office.</span><span class="sxs-lookup"><span data-stu-id="96e70-104">Fluent UI React is the official open-source JavaScript front-end framework designed to build experiences that fit seamlessly into a broad range of Microsoft products, including Office.</span></span> <span data-ttu-id="96e70-105">Ele fornece componentes robustos, atualizados e acessíveis React baseados em React que são altamente personalizáveis usando CSS-in-JS.</span><span class="sxs-lookup"><span data-stu-id="96e70-105">It provides robust, up-to-date, accessible React-based components which are highly customizable using CSS-in-JS.</span></span>

> [!NOTE]
> <span data-ttu-id="96e70-106">Este artigo descreve o uso de interface do usuário fluente React no contexto de Office de complementos. Mas também é usado em uma ampla variedade de Microsoft 365 aplicativos e extensões.</span><span class="sxs-lookup"><span data-stu-id="96e70-106">This article describes the use of Fluent UI React in the context of Office Add-ins. But it is also used in a wide range of Microsoft 365 apps and extensions.</span></span> <span data-ttu-id="96e70-107">Para obter mais informações, [consulte Fluent UI React](https://developer.microsoft.com/fluentui#/get-started/web#fluent-ui-react) and the open source repo Fluent [UI Web](https://github.com/microsoft/fluentui).</span><span class="sxs-lookup"><span data-stu-id="96e70-107">For more information, see [Fluent UI React](https://developer.microsoft.com/fluentui#/get-started/web#fluent-ui-react) and the open source repo [Fluent UI Web](https://github.com/microsoft/fluentui).</span></span>

<span data-ttu-id="96e70-108">Este artigo descreve como criar um complemento criado com o React e usa componentes de interface do usuário React Fluent.</span><span class="sxs-lookup"><span data-stu-id="96e70-108">This article describes how to create an add-in that's built with React and uses Fluent UI React components.</span></span>

## <a name="create-an-add-in-project"></a><span data-ttu-id="96e70-109">Criar um projeto de suplemento</span><span class="sxs-lookup"><span data-stu-id="96e70-109">Create an add-in project</span></span>

<span data-ttu-id="96e70-110">Você usará o gerador Yeoman para Suplementos do Office para criar um projeto de suplemento que use o React.</span><span class="sxs-lookup"><span data-stu-id="96e70-110">You'll use the Yeoman generator for Office Add-ins to create an add-in project that uses React.</span></span>

### <a name="install-the-prerequisites"></a><span data-ttu-id="96e70-111">Instalar os pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="96e70-111">Install the prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-project"></a><span data-ttu-id="96e70-112">Criar o projeto</span><span class="sxs-lookup"><span data-stu-id="96e70-112">Create the project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="96e70-113">**Escolha o tipo de projeto:** `Office Add-in Task Pane project using React framework`</span><span class="sxs-lookup"><span data-stu-id="96e70-113">**Choose a project type:** `Office Add-in Task Pane project using React framework`</span></span>
- <span data-ttu-id="96e70-114">**Escolha o tipo de script:** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="96e70-114">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="96e70-115">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="96e70-115">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="96e70-116">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?**</span><span class="sxs-lookup"><span data-stu-id="96e70-116">**Which Office client application would you like to support?**</span></span> `Word`

![Captura de tela apresentando os avisos e respostas do gerador Yeoman em uma interface de linha de comando](../images/yo-office-word-react.png)

<span data-ttu-id="96e70-118">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="96e70-118">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="try-it-out"></a><span data-ttu-id="96e70-119">Experimente</span><span class="sxs-lookup"><span data-stu-id="96e70-119">Try it out</span></span>

1. <span data-ttu-id="96e70-120">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="96e70-120">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="96e70-121">Conclua as etapas a seguir para iniciar o servidor Web local e fazer o sideload do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="96e70-121">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="96e70-122">Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="96e70-122">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="96e70-123">Se você for solicitado a instalar um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.</span><span class="sxs-lookup"><span data-stu-id="96e70-123">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span> <span data-ttu-id="96e70-124">Você também pode executar o prompt de comando ou terminal como administrador para que as alterações sejam feitas.</span><span class="sxs-lookup"><span data-stu-id="96e70-124">You may also have to run your command prompt or terminal as an administrator for the changes to be made.</span></span>

    > [!TIP]
    > <span data-ttu-id="96e70-125">Se você estiver testando o seu suplemento no Mac, execute o seguinte comando antes de continuar.</span><span class="sxs-lookup"><span data-stu-id="96e70-125">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="96e70-126">O servidor Web local é iniciado quando este comando é executado.</span><span class="sxs-lookup"><span data-stu-id="96e70-126">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="96e70-127">Para testar seu suplemento no Word, execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="96e70-127">To test your add-in in Word, run the following command in the root directory of your project.</span></span> <span data-ttu-id="96e70-128">Isso inicia o servidor Web local (caso ainda não esteja em execução) e abre o Word com o suplemento carregado.</span><span class="sxs-lookup"><span data-stu-id="96e70-128">This starts the local web server (if it's not already running) and opens Word with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="96e70-129">Para testar seu suplemento no Word em um navegador, execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="96e70-129">To test your add-in in Word on a browser, run the following command in the root directory of your project.</span></span> <span data-ttu-id="96e70-130">Quando você executar este comando, o servidor Web local será iniciado (se ainda não estiver em execução).</span><span class="sxs-lookup"><span data-stu-id="96e70-130">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="96e70-131">Para usar o seu suplemento, abra um novo documento no Word na Web e em seguida realize o sideload no suplemento de acordo com as instruções em [Realizar Sideload nos Suplementos do Office no Office na Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="96e70-131">To use your add-in, open a new document in Word on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

3. <span data-ttu-id="96e70-132">Para abrir o painel de tarefas do complemento, na guia **Página** Início, escolha o **botão Mostrar Painel de Tarefas.**</span><span class="sxs-lookup"><span data-stu-id="96e70-132">To open the add-in task pane, on the **Home** tab, choose the **Show Taskpane** button.</span></span> <span data-ttu-id="96e70-133">Observe o texto padrão e o botão **Executar** na parte inferior do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="96e70-133">Notice the default text and the **Run** button at the bottom of the task pane.</span></span> <span data-ttu-id="96e70-134">No restante deste passo a passo, você redefinirá esse texto e um botão criando um componente React que usa componentes de UX do fluent UI React.</span><span class="sxs-lookup"><span data-stu-id="96e70-134">In the remainder of this walkthrough, you'll redefine this text and button by creating a React component that uses UX components from Fluent UI React.</span></span>

    ![Captura de tela mostrando o aplicativo Word com o botão mostrar faixa de opções do Painel de Tarefas realçada e o botão Executar e texto imediatamente anterior realçado no painel de tarefas](../images/word-task-pane-yo-default.png)

## <a name="create-a-react-component-that-uses-fluent-ui-react"></a><span data-ttu-id="96e70-136">Criar um React que usa a interface do usuário do Fluent React</span><span class="sxs-lookup"><span data-stu-id="96e70-136">Create a React component that uses Fluent UI React</span></span>

<span data-ttu-id="96e70-137">Neste ponto, você criou um suplemento muito básico do painel de tarefas usando o React.</span><span class="sxs-lookup"><span data-stu-id="96e70-137">At this point, you've created a very basic task pane add-in that's built using React.</span></span> <span data-ttu-id="96e70-138">Em seguida, siga as etapas abaixo para criar um novo componente React (`ButtonPrimaryExample`) dentro do projeto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="96e70-138">Next, complete the following steps to create a new React component (`ButtonPrimaryExample`) within the add-in project.</span></span> <span data-ttu-id="96e70-139">O componente usa os componentes e da interface `Label` do usuário do Fluent `PrimaryButton` React.</span><span class="sxs-lookup"><span data-stu-id="96e70-139">The component uses the `Label` and `PrimaryButton` components from Fluent UI React.</span></span>

1. <span data-ttu-id="96e70-140">Abra a pasta do projeto criada pelo gerador Yeoman e acesse **src\taskpane\components**.</span><span class="sxs-lookup"><span data-stu-id="96e70-140">Open the project folder created by the Yeoman generator, and go to **src\taskpane\components**.</span></span>
2. <span data-ttu-id="96e70-141">Nesta pasta, crie um novo arquivo chamado **Button.tsx**.</span><span class="sxs-lookup"><span data-stu-id="96e70-141">In that folder, create a new file named **Button.tsx**.</span></span>
3. <span data-ttu-id="96e70-142">Em **Button.tsx**, adicione o código a seguir para definir o componente `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="96e70-142">In **Button.tsx**, add the following code to define the `ButtonPrimaryExample` component.</span></span>

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

<span data-ttu-id="96e70-143">Esse código faz o seguinte:</span><span class="sxs-lookup"><span data-stu-id="96e70-143">This code does the following:</span></span>

- <span data-ttu-id="96e70-144">Faz referência à biblioteca React usando `import * as React from 'react';`.</span><span class="sxs-lookup"><span data-stu-id="96e70-144">References the React library using `import * as React from 'react';`.</span></span>
- <span data-ttu-id="96e70-145">Faz referência aos componentes da interface do usuário do Fluent React ( `PrimaryButton` , , ) que são usados para criar `IButtonProps` `Label` `ButtonPrimaryExample` .</span><span class="sxs-lookup"><span data-stu-id="96e70-145">References the Fluent UI React components (`PrimaryButton`, `IButtonProps`, `Label`) that are used to create `ButtonPrimaryExample`.</span></span>
- <span data-ttu-id="96e70-146">Declara o novo `ButtonPrimaryExample` componente usando `export class ButtonPrimaryExample extends React.Component`.</span><span class="sxs-lookup"><span data-stu-id="96e70-146">Declares the new `ButtonPrimaryExample` component using `export class ButtonPrimaryExample extends React.Component`.</span></span>
- <span data-ttu-id="96e70-147">Declara a `insertText` função que manipulará o evento do `onClick` botão.</span><span class="sxs-lookup"><span data-stu-id="96e70-147">Declares the `insertText` function that will handle the button's `onClick` event.</span></span>
- <span data-ttu-id="96e70-148">Define a interface do usuário do componente do React na função `render`.</span><span class="sxs-lookup"><span data-stu-id="96e70-148">Defines the UI of the React component in the `render` function.</span></span> <span data-ttu-id="96e70-149">A marcação HTML usa os componentes e da interface do usuário do Fluent React especifica que, quando o evento for ativos, a `Label` `PrimaryButton` função será `onClick` `insertText` executado.</span><span class="sxs-lookup"><span data-stu-id="96e70-149">The HTML markup uses the `Label` and `PrimaryButton` components from Fluent UI React and specifies that when the `onClick` event fires, the `insertText` function will run.</span></span>

## <a name="add-the-react-component-to-your-add-in"></a><span data-ttu-id="96e70-150">Adicionar o componente do React ao suplemento</span><span class="sxs-lookup"><span data-stu-id="96e70-150">Add the React component to your add-in</span></span>

<span data-ttu-id="96e70-151">Adicionar o `ButtonPrimaryExample` componente ao suplemento abrindo **src\components\App.tsx** e seguir as seguintes etapas:</span><span class="sxs-lookup"><span data-stu-id="96e70-151">Add the `ButtonPrimaryExample` component to your add-in by opening **src\components\App.tsx** and completing the following steps:</span></span>

1. <span data-ttu-id="96e70-152">Adicione a seguinte declaração de importação para a referência `ButtonPrimaryExample` do **Button.tsx**.</span><span class="sxs-lookup"><span data-stu-id="96e70-152">Add the following import statement to reference `ButtonPrimaryExample` from **Button.tsx**.</span></span>

    ```typescript
    import {ButtonPrimaryExample} from './Button';
    ```

2. <span data-ttu-id="96e70-153">Remova as duas instruções de importação a seguir.</span><span class="sxs-lookup"><span data-stu-id="96e70-153">Remove the following two import statements.</span></span>

    ```typescript
    import { Button, ButtonType } from 'office-ui-fabric-react';
    ...
    import Progress from './Progress';
    ```

3. <span data-ttu-id="96e70-154">Substitua a função padrão `render()` pelo código a seguir, que usa `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="96e70-154">Replace the default `render()` function with the following code that uses `ButtonPrimaryExample`.</span></span>

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

4. <span data-ttu-id="96e70-155">Salve as alterações feitas em **App.tsx**.</span><span class="sxs-lookup"><span data-stu-id="96e70-155">Save the changes you've made to **App.tsx**.</span></span>

## <a name="see-the-result"></a><span data-ttu-id="96e70-156">Ver o resultado</span><span class="sxs-lookup"><span data-stu-id="96e70-156">See the result</span></span>

<span data-ttu-id="96e70-157">No Word, o painel de tarefas do suplemento será atualizado automaticamente quando você salvar as alterações em **App.tsx**.</span><span class="sxs-lookup"><span data-stu-id="96e70-157">In Word, the add-in task pane automatically updates when you save changes to **App.tsx**.</span></span> <span data-ttu-id="96e70-158">O texto padrão e o botão na parte inferior do painel de tarefas agora mostram a IU definida pelo `ButtonPrimaryExample` componente.</span><span class="sxs-lookup"><span data-stu-id="96e70-158">The default text and button at the bottom of the task pane now shows the UI that's defined by the `ButtonPrimaryExample` component.</span></span> <span data-ttu-id="96e70-159">Feche o botão **Insert text...** para inserir o texto no documento.</span><span class="sxs-lookup"><span data-stu-id="96e70-159">Choose the **Insert text...** button to insert text into the document.</span></span>

![Captura de tela mostrando o aplicativo Word com o "Inserir texto..." botão e texto imediatamente anterior realçado](../images/word-task-pane-with-react-component.png)

<span data-ttu-id="96e70-161">Parabéns, você criou com êxito um complemento do painel de tarefas usando o React e a interface do usuário React!</span><span class="sxs-lookup"><span data-stu-id="96e70-161">Congratulations, you've successfully created a task pane add-in using React and Fluent UI React!</span></span>

## <a name="see-also"></a><span data-ttu-id="96e70-162">Confira também</span><span class="sxs-lookup"><span data-stu-id="96e70-162">See also</span></span>

- [<span data-ttu-id="96e70-163">Word Add-in GettingStartedFabricReact</span><span class="sxs-lookup"><span data-stu-id="96e70-163">Word Add-in GettingStartedFabricReact</span></span>](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [<span data-ttu-id="96e70-164">Fabric Core em Office de complementos</span><span class="sxs-lookup"><span data-stu-id="96e70-164">Fabric Core in Office Add-ins</span></span>](fabric-core.md)
- [<span data-ttu-id="96e70-165">Padrões de design da experiência do usuário para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="96e70-165">UX design patterns for Office Add-ins</span></span>](ux-design-pattern-templates.md)
