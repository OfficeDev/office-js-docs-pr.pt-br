---
ms.date: 07/07/2020
ms.prod: non-product-specific
description: Tutorial sobre como compartilhar código entre um suplemento VSTO e um suplemento do Office.
title: 'Tutorial: compartilhar código entre um suplemento VSTO e um suplemento do Office usando uma biblioteca de códigos compartilhado'
localization_priority: Priority
ms.openlocfilehash: 761820dece1d5b8322de38863e10ad2f536445b9
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49131742"
---
# <a name="tutorial-share-code-between-both-a-vsto-add-in-and-an-office-add-in-with-a-shared-code-library"></a><span data-ttu-id="73915-103">Tutorial: compartilhar código entre um suplemento VSTO e um suplemento do Office com uma biblioteca de códigos compartilhadas</span><span class="sxs-lookup"><span data-stu-id="73915-103">Tutorial: Share code between both a VSTO Add-in and an Office add-in with a shared code library</span></span>

<span data-ttu-id="73915-104">Os suplementos do Visual Studio Tools for Office (VSTO) são ótimos para a ampliação do Office para fornecer soluções para seus negócios ou para outras pessoas.</span><span class="sxs-lookup"><span data-stu-id="73915-104">Visual Studio Tools for Office (VSTO) Add-ins are great for extending Office to provide solutions for your business or others.</span></span> <span data-ttu-id="73915-105">Eles já estão por aqui há muito tempo e há milhares de soluções criadas com o VSTO.</span><span class="sxs-lookup"><span data-stu-id="73915-105">They've been around for a long time and there are thousands of solutions built with VSTO.</span></span> <span data-ttu-id="73915-106">No entanto, eles só são executados no Office no Windows.</span><span class="sxs-lookup"><span data-stu-id="73915-106">However, they only run on Office on Windows.</span></span> <span data-ttu-id="73915-107">Não é possível executar suplementos VSTO no Mac, online ou em plataformas móveis.</span><span class="sxs-lookup"><span data-stu-id="73915-107">You can't run VSTO Add-ins on Mac, online, or mobile platforms.</span></span>

<span data-ttu-id="73915-108">Os suplementos do Office usam HTML, JavaScript e tecnologias da Web adicionais para criar soluções do Office em todas as plataformas.</span><span class="sxs-lookup"><span data-stu-id="73915-108">Office Add-ins use HTML, JavaScript, and additional web technologies to build Office solutions on all platforms.</span></span> <span data-ttu-id="73915-109">Migrar seu suplemento existente do VSTO para um suplemento do Office é uma ótima maneira de disponibilizá-lo em todas as plataformas.</span><span class="sxs-lookup"><span data-stu-id="73915-109">Migrating your existing VSTO Add-in to an Office add-in is a great way to make your solution available across all platforms.</span></span>

<span data-ttu-id="73915-110">Talvez você queira manter o suplemento VSTO e um novo suplemento do Office que tenham a mesma funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="73915-110">You may want to maintain both your VSTO Add-in and a new Office add-in that both have the same functionality.</span></span> <span data-ttu-id="73915-111">Isso permite que você continue servindo aos clientes que usam o suplemento VSTO no Office no Windows.</span><span class="sxs-lookup"><span data-stu-id="73915-111">This enables you to continue servicing your customers that use the VSTO Add-in on Office on Windows.</span></span> <span data-ttu-id="73915-112">Isso também permite fornecer a mesma funcionalidade em um suplemento do Office para clientes em todas as plataformas.</span><span class="sxs-lookup"><span data-stu-id="73915-112">This also enables you to provide the same functionality in an Office add-in for customers across all platforms.</span></span> <span data-ttu-id="73915-113">Você também pode [tornar seu suplemento do Office compatível com o suplemento VSTO existente](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="73915-113">You can also [Make your Office add-in compatible with the existing VSTO Add-in](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).</span></span>

<span data-ttu-id="73915-114">No entanto, é melhor evitar a reconfiguração de todo o código de seu suplemento VSTO para o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="73915-114">However it is best to avoid rewriting all the code from your VSTO Add-in for the Office add-in.</span></span> <span data-ttu-id="73915-115">Este tutorial mostra como evitar a reconfiguração de código usando uma biblioteca compartilhadas de códigos para ambos os suplementos.</span><span class="sxs-lookup"><span data-stu-id="73915-115">This tutorial shows how to avoid rewriting code by using a shared code library for both add-ins.</span></span>

## <a name="shared-code-library"></a><span data-ttu-id="73915-116">Biblioteca de códigos compartilhados</span><span class="sxs-lookup"><span data-stu-id="73915-116">Shared code library</span></span>

<span data-ttu-id="73915-117">Este tutorial orientará você pelas etapas de identificação e compartilhamento de códigos comuns entre seu suplemento VSTO e um suplemento moderno do Office.</span><span class="sxs-lookup"><span data-stu-id="73915-117">This tutorial will walk you through the steps of identifying and sharing common code between your VSTO Add-in and a modern Office add-in.</span></span> <span data-ttu-id="73915-118">Ele usa um exemplo de suplemento VSTO muito simples para as etapas para que você possa se concentrar nas habilidades e técnicas necessárias para trabalhar com seus próprios suplementos do VSTO.</span><span class="sxs-lookup"><span data-stu-id="73915-118">It uses a very simple VSTO Add-in example for the steps so that you can focus on the skills and techniques you will need for working with your own VSTO Add-ins.</span></span>

<span data-ttu-id="73915-119">O diagrama a seguir mostra como a biblioteca de códigos compartilhada funciona para migração.</span><span class="sxs-lookup"><span data-stu-id="73915-119">The following diagram shows how the shared code library works for migration.</span></span> <span data-ttu-id="73915-120">O código comum é refatorado em uma nova biblioteca de códigos compartilhadas.</span><span class="sxs-lookup"><span data-stu-id="73915-120">Common code is refactored into a new shared code library.</span></span> <span data-ttu-id="73915-121">O código pode permanecer escrito em seu idioma original, como o C# ou o VB.</span><span class="sxs-lookup"><span data-stu-id="73915-121">The code can remain written in its original language, such as C# or VB.</span></span> <span data-ttu-id="73915-122">Isso significa que você pode continuar usando o código do suplemento VSTO existente, criando uma referência do projeto.</span><span class="sxs-lookup"><span data-stu-id="73915-122">This means you can continue using the code in the existing VSTO Add-in by creating a project reference.</span></span> <span data-ttu-id="73915-123">Quando você cria o suplemento do Office, ele também usa a biblioteca compartilhadas de códigos chamando-a por APIs REST.</span><span class="sxs-lookup"><span data-stu-id="73915-123">When you create the Office add-in, it will also use the shared code library by calling into it through REST APIs.</span></span>

![Diagrama de suplemento VSTO e suplemento do Office usando uma biblioteca de códigos compartilhados](../images/vsto-migration-shared-code-library.png)

<span data-ttu-id="73915-125">Habilidades e técnicas neste tutorial:</span><span class="sxs-lookup"><span data-stu-id="73915-125">Skills and techniques in this tutorial:</span></span>

- <span data-ttu-id="73915-126">Criar uma biblioteca de classe compartilhada, refatorando o código em uma biblioteca de classe do .NET.</span><span class="sxs-lookup"><span data-stu-id="73915-126">Create a shared class library by refactoring code into a .NET class library.</span></span>
- <span data-ttu-id="73915-127">Crie um invólucro da API REST usando ASP.NET Core para a biblioteca de classe compartilhada.</span><span class="sxs-lookup"><span data-stu-id="73915-127">Create a REST API wrapper using ASP.NET Core for the shared class library.</span></span>
- <span data-ttu-id="73915-128">Chame a API REST do suplemento do Office para acessar o código compartilhado.</span><span class="sxs-lookup"><span data-stu-id="73915-128">Call the REST API from the Office add-in to access shared code.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="73915-129">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="73915-129">Prerequisites</span></span>

<span data-ttu-id="73915-130">Para configurar seu ambiente de desenvolvimento:</span><span class="sxs-lookup"><span data-stu-id="73915-130">To set up your development environment:</span></span>

1. <span data-ttu-id="73915-131">Instalar o [Visual Studio 2019](https://visualstudio.microsoft.com/downloads/).</span><span class="sxs-lookup"><span data-stu-id="73915-131">Install [Visual Studio 2019](https://visualstudio.microsoft.com/downloads/).</span></span>
2. <span data-ttu-id="73915-132">Instalar as seguintes cargas de trabalho:</span><span class="sxs-lookup"><span data-stu-id="73915-132">Install the following workloads:</span></span>
    - <span data-ttu-id="73915-133">ASP.NET e desenvolvimento na Web</span><span class="sxs-lookup"><span data-stu-id="73915-133">ASP.NET and web development</span></span>
    - <span data-ttu-id="73915-134">Desenvolvimento de várias plataformas do .NET Core.</span><span class="sxs-lookup"><span data-stu-id="73915-134">.NET Core cross-platform development.</span></span>
    - <span data-ttu-id="73915-135">Desenvolvimento do Office/SharePoint</span><span class="sxs-lookup"><span data-stu-id="73915-135">Office/SharePoint development</span></span>
    - <span data-ttu-id="73915-136">Os seguintes componentes **individuais**.</span><span class="sxs-lookup"><span data-stu-id="73915-136">The following **Individual** components.</span></span>
        - <span data-ttu-id="73915-137">Ferramentas do Visual Studio para Office (VSTO)</span><span class="sxs-lookup"><span data-stu-id="73915-137">Visual Studio Tools for Office (VSTO).</span></span>
        - <span data-ttu-id="73915-138">.NET Core 3.0 Runtime.</span><span class="sxs-lookup"><span data-stu-id="73915-138">.NET Core 3.0 Runtime.</span></span>

<span data-ttu-id="73915-139">Também são necessários:</span><span class="sxs-lookup"><span data-stu-id="73915-139">You also need the following:</span></span>

- <span data-ttu-id="73915-140">Uma conta do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="73915-140">A Microsoft 365 account.</span></span> <span data-ttu-id="73915-141">Você pode se cadastrar no [programa de desenvolvedores do Microsoft 365](https://aka.ms/devprogramsignup), que inclui um ano de assinatura gratuita do Office 365.</span><span class="sxs-lookup"><span data-stu-id="73915-141">You can join the [Microsoft 365 developer program](https://aka.ms/devprogramsignup) that includes a free 1 year subscription to Office 365.</span></span>
- <span data-ttu-id="73915-142">Um Locatário do Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="73915-142">A Microsoft Azure Tenant.</span></span> <span data-ttu-id="73915-143">Você pode adquirir uma assinatura de avaliação no [Microsoft Azure](https://account.windowsazure.com/SignUp).</span><span class="sxs-lookup"><span data-stu-id="73915-143">A trial subscription can be acquired here: [Microsoft Azure](https://account.windowsazure.com/SignUp).</span></span>

## <a name="the-cell-analyzer-vsto-add-in"></a><span data-ttu-id="73915-144">O suplemento VSTO do analisador de células</span><span class="sxs-lookup"><span data-stu-id="73915-144">The Cell analyzer VSTO Add-in</span></span>

<span data-ttu-id="73915-145">Este tutorial usa a solução PnP [Biblioteca compartilhada do suplemento VSTO para o suplemento do Office](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/VSTO-shared-code-migration).</span><span class="sxs-lookup"><span data-stu-id="73915-145">This tutorial uses the [VSTO Add-in shared library for Office add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/VSTO-shared-code-migration) PnP solution.</span></span> <span data-ttu-id="73915-146">A pasta **/start** contém a solução de suplemento VSTO que você migrará.</span><span class="sxs-lookup"><span data-stu-id="73915-146">The **/start** folder contains the VSTO Add-in solution that you will migrate.</span></span> <span data-ttu-id="73915-147">Sua meta é migrar o suplemento VSTO para um suplemento moderno do Office, quando possível.</span><span class="sxs-lookup"><span data-stu-id="73915-147">Your goal is to migrate the VSTO Add-in to a modern Office add-in by sharing code when possible.</span></span>

> [!NOTE]
> <span data-ttu-id="73915-148">O exemplo usa C#, mas você pode aplicar as técnicas deste tutorial a um suplemento VSTO escrito em qualquer linguagem .NET.</span><span class="sxs-lookup"><span data-stu-id="73915-148">The sample uses C# but you can apply the techniques in this tutorial to a VSTO Add-in written in any .NET language.</span></span>

1. <span data-ttu-id="73915-149">Baixe a solução PnP [Biblioteca compartilhada do suplemento VSTO para o suplemento do Office](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/VSTO-shared-code-migration)para trabalhar em um arquivo em seu computador.</span><span class="sxs-lookup"><span data-stu-id="73915-149">Download the [VSTO Add-in shared library for Office add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/VSTO-shared-code-migration) PnP solution to a working folder on your computer.</span></span>
1. <span data-ttu-id="73915-150">Inicie o Visual Studio 2019 e abra a solução **/start/Cell-Analyzer.sln**.</span><span class="sxs-lookup"><span data-stu-id="73915-150">Start Visual Studio 2019 and open the **/start/Cell-Analyzer.sln** solution.</span></span>
1. <span data-ttu-id="73915-151">No menu **Depurar**, selecione **Iniciar Depuração**.</span><span class="sxs-lookup"><span data-stu-id="73915-151">On the **Debug** menu, choose **Start Debugging**.</span></span>
1. <span data-ttu-id="73915-152">No **Gerenciador de soluções**, clique com o botão direito do mouse no projeto **Cell-Analyzer** e escolha **Propriedades**.</span><span class="sxs-lookup"><span data-stu-id="73915-152">In **Solution Explorer**, right-click the **Cell-Analyzer** project, and choose **Properties**.</span></span>
1. <span data-ttu-id="73915-153">Escolha a categoria **Assinatura** nas propriedades.</span><span class="sxs-lookup"><span data-stu-id="73915-153">Choose the **Signing** category in the properties.</span></span>
1. <span data-ttu-id="73915-154">Escolha **Assinar os manifestos ClickOnce** e, em seguida, escolha **Criar certificado de teste**.</span><span class="sxs-lookup"><span data-stu-id="73915-154">Choose **Sign the ClickOnce manifests**, and then chose **Create Test Certificate**.</span></span>
1. <span data-ttu-id="73915-155">Na caixa de diálogo **criar certificado de teste**, digite e confirme a senha.</span><span class="sxs-lookup"><span data-stu-id="73915-155">In the **Create Test Certificate** dialog, enter and confirm a password.</span></span> <span data-ttu-id="73915-156">Em seguida, escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="73915-156">Then choose **OK**.</span></span>

<span data-ttu-id="73915-157">O suplemento é um painel de tarefas personalizado do Excel.</span><span class="sxs-lookup"><span data-stu-id="73915-157">The add-in is a custom task pane for Excel.</span></span> <span data-ttu-id="73915-158">Você pode selecionar qualquer célula com o texto e escolher o botão **Mostrar o Unicode**.</span><span class="sxs-lookup"><span data-stu-id="73915-158">You can select any cell with text, and then choose the **Show unicode** button.</span></span> <span data-ttu-id="73915-159">Na seção **Resultado**, o suplemento exibirá uma lista de cada caractere no texto junto com seu número Unicode correspondente.</span><span class="sxs-lookup"><span data-stu-id="73915-159">In the **Result** section, the add-in will display a list of each character in the text along with its corresponding Unicode number.</span></span>

![Captura de tela do suplemento VSTO do Analisador de células executado em Excel com o botão "Mostrar Unicode" e Seção de resultados vazia](../images/pnp-cell-analyzer-vsto-add-in.png)

## <a name="analyze-types-of-code-in-the-vsto-add-in"></a><span data-ttu-id="73915-161">Análise dos tipos de código no Suplemento VSTO</span><span class="sxs-lookup"><span data-stu-id="73915-161">Analyze types of code in the VSTO Add-in</span></span>

<span data-ttu-id="73915-162">A primeira técnica a ser aplicada é analisar o suplemento para quais partes do código podem ser compartilhadas.</span><span class="sxs-lookup"><span data-stu-id="73915-162">The first technique to apply is to analyze the add-in for which parts of code can be shared.</span></span> <span data-ttu-id="73915-163">Em geral, o Project é dividido em três tipos de códigos.</span><span class="sxs-lookup"><span data-stu-id="73915-163">In general, project will break down into three types of code.</span></span>

### <a name="ui-code"></a><span data-ttu-id="73915-164">Código IU</span><span class="sxs-lookup"><span data-stu-id="73915-164">UI code</span></span>

<span data-ttu-id="73915-165">O código da IU interage com o usuário.</span><span class="sxs-lookup"><span data-stu-id="73915-165">UI code interacts with the user.</span></span> <span data-ttu-id="73915-166">O código da interface de usuário do VSTO funciona com formulários do Windows.</span><span class="sxs-lookup"><span data-stu-id="73915-166">In VSTO UI code works through Windows Forms.</span></span> <span data-ttu-id="73915-167">Os suplementos do Office usam HTML, CSS e JavaScript para IU.</span><span class="sxs-lookup"><span data-stu-id="73915-167">Office Add-ins use HTML, CSS, and JavaScript for UI.</span></span> <span data-ttu-id="73915-168">Devido a essas diferenças, não é possível compartilhar o código da interface do usuário com o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="73915-168">Because of these differences you cannot share UI code to the Office add-in.</span></span> <span data-ttu-id="73915-169">A IU deve ser recriada em JavaScript.</span><span class="sxs-lookup"><span data-stu-id="73915-169">UI will need to be recreated in JavaScript.</span></span>

### <a name="document-code"></a><span data-ttu-id="73915-170">Código do documento</span><span class="sxs-lookup"><span data-stu-id="73915-170">Document code</span></span>

<span data-ttu-id="73915-171">O código VSTO interage com o documento por meio de objetos .NET, como `Microsoft.Office.Interop.Excel.Range`.</span><span class="sxs-lookup"><span data-stu-id="73915-171">In VSTO code interacts with the document through .NET objects such as `Microsoft.Office.Interop.Excel.Range`.</span></span> <span data-ttu-id="73915-172">No entanto, os suplementos do Office usam a biblioteca Office.js.</span><span class="sxs-lookup"><span data-stu-id="73915-172">But Office Add-ins use the Office.js library.</span></span> <span data-ttu-id="73915-173">Embora sejam similares, eles não são exatamente iguais.</span><span class="sxs-lookup"><span data-stu-id="73915-173">Although these are similar, they are not exactly the same.</span></span> <span data-ttu-id="73915-174">Portanto, você não pode compartilhar o código de interação do documento com o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="73915-174">So again, you cannot share document interaction code to the Office add-in.</span></span>

### <a name="logic-code"></a><span data-ttu-id="73915-175">Código lógico</span><span class="sxs-lookup"><span data-stu-id="73915-175">Logic code</span></span>

<span data-ttu-id="73915-176">A lógica empresarial, algoritmos, funções auxiliares e um código semelhante geralmente formam o coração de um suplemento VSTO.</span><span class="sxs-lookup"><span data-stu-id="73915-176">Business logic, algorithms, helper functions, and similar code often make up the heart of a VSTO Add-in.</span></span> <span data-ttu-id="73915-177">Esse código funciona independentemente da interface de usuário e do código do documento para executar a análise, conectar-se a serviços de backend, executar cálculos e muito mais.</span><span class="sxs-lookup"><span data-stu-id="73915-177">This code works independently of the UI and document code to perform analysis, connect to backend services, run calculations, and more.</span></span> <span data-ttu-id="73915-178">Esse é o código que pode ser compartilhado para que você não precise escrevê-lo novamente em JavaScript.</span><span class="sxs-lookup"><span data-stu-id="73915-178">This is the code that can be shared so that you don't have to rewrite it in JavaScript.</span></span>

<span data-ttu-id="73915-179">Vamos examinar o suplemento VSTO.</span><span class="sxs-lookup"><span data-stu-id="73915-179">Let's examine the VSTO Add-in.</span></span> <span data-ttu-id="73915-180">No código a seguir, cada seção é identificada como um código de documento, IU ou de algoritmo.</span><span class="sxs-lookup"><span data-stu-id="73915-180">In the following code, each section is identified as DOCUMENT, UI, or ALGORITHM code.</span></span>

```csharp
// **_ UI CODE _*_
private void btnUnicode_Click(object sender, EventArgs e)
{
    // _*_ DOCUMENT CODE _*_
    Microsoft.Office.Interop.Excel.Range rangeCell;
    rangeCell = Globals.ThisAddIn.Application.ActiveCell;

    string cellValue = "";

    if (null != rangeCell.Value)
    {
        cellValue = rangeCell.Value.ToString();
    }

    // _*_ ALGORITHM CODE _*_
    //convert string to Unicode listing
    string result = "";
    foreach (char c in cellValue)
    {
        int unicode = c;

        result += $"{c}: {unicode}\r\n";
    }

    // _*_ UI CODE _*_
    //Output the result
    txtResult.Text = result;
}
```

<span data-ttu-id="73915-181">Com essa abordagem, você pode ver que uma seção de código pode ser compartilhada com o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="73915-181">Using this approach you can see that one section of code can be shared to the Office add-in.</span></span> <span data-ttu-id="73915-182">O código a seguir precisará ser refatorado em uma biblioteca de classe separada.</span><span class="sxs-lookup"><span data-stu-id="73915-182">The following code will need to be refactored into a separate class library.</span></span>

```csharp
// _*_ ALGORITHM CODE _*_
//convert string to Unicode listing
string result = "";
foreach (char c in cellValue)
{
    int unicode = c;

    result += $"{c}: {unicode}\r\n";
}
```

## <a name="create-a-shared-class-library"></a><span data-ttu-id="73915-183">Criar uma biblioteca de classe compartilhada</span><span class="sxs-lookup"><span data-stu-id="73915-183">Create a shared class library</span></span>

<span data-ttu-id="73915-184">Os suplementos do VSTO são criados no Visual Studio como projetos .NET, portanto, reutilizaremos o .NET o máximo possível para simplificar.</span><span class="sxs-lookup"><span data-stu-id="73915-184">VSTO Add-ins are created in Visual Studio as .NET projects, so we'll reuse .NET as much as possible to keep things simple.</span></span> <span data-ttu-id="73915-185">Nossa próxima técnica é criar uma biblioteca de classe e um código compartilhado de refatoração nessa biblioteca de classe.</span><span class="sxs-lookup"><span data-stu-id="73915-185">Our next technique is to create a class library and refactor shared code into that class library.</span></span>

1. <span data-ttu-id="73915-186">Caso você ainda não o tenha feito, inicie o Visual Studio 2019 e abra a solução \*\start\Cell-Analyzer.sln\*\*.</span><span class="sxs-lookup"><span data-stu-id="73915-186">If you haven't already, start Visual Studio 2019 and open the _ *\start\Cell-Analyzer.sln*\* solution.</span></span>
2. <span data-ttu-id="73915-187">Clique com botão direito do mouse da solução em **Gerenciador de soluções** e escolha **Adicionar > Novo projeto**.</span><span class="sxs-lookup"><span data-stu-id="73915-187">Right-click the solution in **Solution Explorer** and choose **Add > New Project**.</span></span>
3. <span data-ttu-id="73915-188">Na caixa de diálogo **Adicionar um novo projeto**, escolha **Biblioteca de Classe (.NET Framework)** e escolha **Próximo**.</span><span class="sxs-lookup"><span data-stu-id="73915-188">In the **Add a new project dialog**, choose **Class Library (.NET Framework)**, and choose **Next**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="73915-189">Não use a biblioteca de classe central do .NET porque ela não funcionará com seu projeto do VSTO.</span><span class="sxs-lookup"><span data-stu-id="73915-189">Don't use the .NET Core class library because it will not work with your VSTO project.</span></span>
4. <span data-ttu-id="73915-190">Na caixa de diálogo **Configure seu novo Project**, defina os seguintes campos.</span><span class="sxs-lookup"><span data-stu-id="73915-190">In the **Configure your new project** dialog, set the following fields.</span></span>
    - <span data-ttu-id="73915-191">Defina o **Nome do projeto** como **CellAnalyzerSharedLibrary**.</span><span class="sxs-lookup"><span data-stu-id="73915-191">Set the **Project name** to **CellAnalyzerSharedLibrary**.</span></span>
    - <span data-ttu-id="73915-192">Deixe o **Local** com o valor padrão.</span><span class="sxs-lookup"><span data-stu-id="73915-192">Leave the **Location** at it's default value.</span></span>
    - <span data-ttu-id="73915-193">Defina a **estrutura** como **4.7.2**.</span><span class="sxs-lookup"><span data-stu-id="73915-193">Set the **Framework** to **4.7.2**.</span></span>
5. <span data-ttu-id="73915-194">Escolha **Criar**.</span><span class="sxs-lookup"><span data-stu-id="73915-194">Choose **Create**.</span></span>
6. <span data-ttu-id="73915-195">Depois de criar o projeto, renomeie o arquivo **Class1.cs** para **CellOperations.cs**.</span><span class="sxs-lookup"><span data-stu-id="73915-195">After the project is created, rename the **Class1.cs** file to **CellOperations.cs**.</span></span> <span data-ttu-id="73915-196">Será exibida uma solicitação para renomear a classe.</span><span class="sxs-lookup"><span data-stu-id="73915-196">A prompt to rename the class appears.</span></span> <span data-ttu-id="73915-197">Renomeie o nome da classe para que ele corresponda ao nome do arquivo.</span><span class="sxs-lookup"><span data-stu-id="73915-197">Rename the class name so that it matches the file name.</span></span>
7. <span data-ttu-id="73915-198">Adicione o seguinte código à classe `CellOperations` para criar um método chamado `GetUnicodeFromText`.</span><span class="sxs-lookup"><span data-stu-id="73915-198">Add the following code to the `CellOperations` class to create a method named `GetUnicodeFromText`.</span></span>

```csharp
public class CellOperations
{
    static public string GetUnicodeFromText(string value)
    {
        string result = "";
        foreach (char c in value)
        {
            int unicode = c;

            result += $"{c}: {unicode}\r\n";
        }
        return result;
    }
}
```

### <a name="use-the-shared-class-library-in-the-vsto-add-in"></a><span data-ttu-id="73915-199">Use a biblioteca de classe compartilhada no suplemento VSTO</span><span class="sxs-lookup"><span data-stu-id="73915-199">Use the shared class library in the VSTO Add-in</span></span>

<span data-ttu-id="73915-200">Agora, você precisa atualizar o suplemento VSTO para usar a biblioteca de classe.</span><span class="sxs-lookup"><span data-stu-id="73915-200">Now you need to update the VSTO Add-in to use the class library.</span></span> <span data-ttu-id="73915-201">É importante que o suplemento VSTO e o suplemento do Office usem a mesma biblioteca de classes compartilhadas para que correções de bugs futuras ou recursos sejam feitos em um único local.</span><span class="sxs-lookup"><span data-stu-id="73915-201">This is important that both the VSTO Add-in and Office add-in use the same shared class library so that future bug fixes or features are made in one location.</span></span>

1. <span data-ttu-id="73915-202">No **Gerenciador de soluções**, clique com o botão direito do mouse em **Cell-Analyzer** e escolha **Adicionar referência**.</span><span class="sxs-lookup"><span data-stu-id="73915-202">In **Solution Explorer** right-click the **Cell-Analyzer** project, and choose **Add Reference**.</span></span>
2. <span data-ttu-id="73915-203">Selecione **CellAnalyzerSharedLibrary** e escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="73915-203">Select **CellAnalyzerSharedLibrary**, and choose **OK**.</span></span>
3. <span data-ttu-id="73915-204">No **Gerenciador de soluções** expanda o arquivo **Cell-Analyzer**, clique com o botão direito do mouse no arquivo **CellAnalyzerPane.cs** e escolha **Exibir Código**.</span><span class="sxs-lookup"><span data-stu-id="73915-204">In **Solution Explorer** expand the **Cell-Analyzer** project, right-click the **CellAnalyzerPane.cs** file, and choose **View Code**.</span></span>
4. <span data-ttu-id="73915-205">No método `btnUnicode_Click`, exclua as linhas de código a seguir.</span><span class="sxs-lookup"><span data-stu-id="73915-205">In the `btnUnicode_Click` method, delete the following lines of code.</span></span>

    ```csharp
    //Convert to Unicode listing
    string result = "";
    foreach (char c in cellValue)
    {
      int unicode = c;
      result += $"{c}: {unicode}\r\n";
    }
    ```

5. <span data-ttu-id="73915-206">Atualize a linha de código sob o comentário `//Output the result` para ler da seguinte maneira:</span><span class="sxs-lookup"><span data-stu-id="73915-206">Update the line of code under the `//Output the result` comment to read as follows:</span></span>

    ```csharp
    //Output the result
    txtResult.Text = CellAnalyzerSharedLibrary.CellOperations.GetUnicodeFromText(cellValue);
    ```

6. <span data-ttu-id="73915-207">No menu **Depurar**, selecione **Iniciar Depuração**.</span><span class="sxs-lookup"><span data-stu-id="73915-207">On the **Debug** menu, choose **Start Debugging**.</span></span> <span data-ttu-id="73915-208">O painel de tarefas personalizado deve funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="73915-208">The custom task pane should work as expected.</span></span> <span data-ttu-id="73915-209">Digite um texto em uma célula e, em seguida, teste para convertê-lo em uma lista Unicode com o suplemento.</span><span class="sxs-lookup"><span data-stu-id="73915-209">Enter some text in a cell, and then test that you can convert it to a Unicode list with the add-in.</span></span>

## <a name="create-a-rest-api-wrapper"></a><span data-ttu-id="73915-210">Criar um invólucro da API REST</span><span class="sxs-lookup"><span data-stu-id="73915-210">Create a REST API wrapper</span></span>

<span data-ttu-id="73915-211">O suplemento VSTO pode usar a biblioteca de classes compartilhadas diretamente, uma vez que ambos são projetos .NET.</span><span class="sxs-lookup"><span data-stu-id="73915-211">The VSTO Add-in can use the shared class library directly since they are both .NET projects.</span></span> <span data-ttu-id="73915-212">No entanto, o suplemento do Office não poderá usar o .NET, uma vez que ele usa o JavaScript.</span><span class="sxs-lookup"><span data-stu-id="73915-212">However the Office add-in won't be able to use .NET since it uses JavaScript.</span></span> <span data-ttu-id="73915-213">Em seguida, você precisará criar um invólucro da API REST.</span><span class="sxs-lookup"><span data-stu-id="73915-213">Next you will need to create a REST API wrapper.</span></span> <span data-ttu-id="73915-214">Isso permite que o suplemento do Office chame uma API REST, que passa a chamada para a biblioteca de classes compartilhadas.</span><span class="sxs-lookup"><span data-stu-id="73915-214">This enables the Office add-in to call a REST API, which then passes the call along to the shared class library.</span></span>

1. <span data-ttu-id="73915-215">No **Gerenciador de soluções**, clique com o botão direito do mouse no **Cell-Analyzer** e escolha **Adicionar > Novo Projeto**.</span><span class="sxs-lookup"><span data-stu-id="73915-215">In **Solution Explorer**, right-click the **Cell-Analyzer** project, and choose **Add > New Project**.</span></span>
2. <span data-ttu-id="73915-216">Em **Adicionar uma nova caixa de diálogo do projeto**, escolha **Aplicativo Web ASP.NET Core** e escolha **Próximo**.</span><span class="sxs-lookup"><span data-stu-id="73915-216">In the **Add a new project dialog**, choose **ASP.NET Core Web Application**, and choose **Next**.</span></span>
3. <span data-ttu-id="73915-217">Na caixa de diálogo **Configure seu novo projeto**, defina os seguintes campos:</span><span class="sxs-lookup"><span data-stu-id="73915-217">In the **Configure your new project** dialog, set the following fields:</span></span>
    - <span data-ttu-id="73915-218">Defina o **nome do projeto** para **CellAnalyzerRESTAPI**.</span><span class="sxs-lookup"><span data-stu-id="73915-218">Set the **Project name** to **CellAnalyzerRESTAPI**.</span></span>
    - <span data-ttu-id="73915-219">No campo **Local**, deixe o valor padrão.</span><span class="sxs-lookup"><span data-stu-id="73915-219">In the **Location** field, leave the default value.</span></span>
4. <span data-ttu-id="73915-220">Escolha **Criar**.</span><span class="sxs-lookup"><span data-stu-id="73915-220">Choose **Create**.</span></span>
5. <span data-ttu-id="73915-221">Na caixa de diálogo **criar um novo aplicativo Web ASP.NET Core**, selecione **ASP.NET Core 3.1** da versão e selecione **API** na lista de projetos.</span><span class="sxs-lookup"><span data-stu-id="73915-221">In the **Create a new ASP.NET Core web application** dialog, select **ASP.NET Core 3.1** for the version, and select **API** in the list of projects.</span></span>
6. <span data-ttu-id="73915-222">Deixe todos os outros campos em valores padrão e escolha o botão **Criar**.</span><span class="sxs-lookup"><span data-stu-id="73915-222">Leave all other fields at default values and choose the **Create** button.</span></span>
7. <span data-ttu-id="73915-223">Depois de criar o projeto, expanda o projeto **CellAnalyzerRESTAPI** no **Gerenciador de soluções**.</span><span class="sxs-lookup"><span data-stu-id="73915-223">After the project is created, expand the **CellAnalyzerRESTAPI** project in **Solution Explorer**.</span></span>
8. <span data-ttu-id="73915-224">Clique com o botão direito do mouse em **Dependências** e escolha **Adicionar Referência**.</span><span class="sxs-lookup"><span data-stu-id="73915-224">Right-click **Dependencies**, and choose **Add Reference**.</span></span>
9. <span data-ttu-id="73915-225">Selecione **CellAnalyzerSharedLibrary** e escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="73915-225">Select **CellAnalyzerSharedLibrary**, and choose **OK**.</span></span>
10. <span data-ttu-id="73915-226">Clique com o botão direito do mouse na pasta **Controladores** e escolha **Adicionar > Controlador**.</span><span class="sxs-lookup"><span data-stu-id="73915-226">Right-click the **Controllers** folder, and choose **Add > Controller**.</span></span>
11. <span data-ttu-id="73915-227">Na caixa de diálogo **Adicionar Novo Item de Scaffolded**, escolha **controlador da API-vazio** e **Adicionar**.</span><span class="sxs-lookup"><span data-stu-id="73915-227">In the **Add New Scaffolded Item** dialog, choose **API Controller - Empty** and then **Add**.</span></span>
12. <span data-ttu-id="73915-228">Na caixa de diálogo **Adicionar controlador de API vazio**, nomeie o controlador **AnalyzeUnicodeController** e selecione **Adicionar**.</span><span class="sxs-lookup"><span data-stu-id="73915-228">In the **Add Empty API Controller** dialog, name the controller **AnalyzeUnicodeController**, and then choose **Add**.</span></span>
13. <span data-ttu-id="73915-229">Abra o arquivo **AnalyzeUnicodeController.cs** e adicione o código a seguir como um método para a classe `AnalyzeUnicodeController`.</span><span class="sxs-lookup"><span data-stu-id="73915-229">Open the **AnalyzeUnicodeController.cs** file and add the following code as a method to the `AnalyzeUnicodeController` class.</span></span>

    ```csharp
    [HttpGet]
    public ActionResult<string> AnalyzeUnicode(string value)
    {
      if (value == null)
      {
        return BadRequest();
      }
      return CellAnalyzerSharedLibrary.CellOperations.GetUnicodeFromText(value);
    }
    ```

14. <span data-ttu-id="73915-230">Clique com o botão direito do mouse no projeto **CellAnalyzerRESTAPI** e escolha **Definir como inicialização do projeto**.</span><span class="sxs-lookup"><span data-stu-id="73915-230">Right-click the **CellAnalyzerRESTAPI** project, and choose **Set as Startup Project**.</span></span>
15. <span data-ttu-id="73915-231">No menu **Depurar**, selecione **Iniciar Depuração**.</span><span class="sxs-lookup"><span data-stu-id="73915-231">On the **Debug** menu, choose **Start Debugging**.</span></span>
16. <span data-ttu-id="73915-232">Um navegador será iniciado.</span><span class="sxs-lookup"><span data-stu-id="73915-232">A browser will launch.</span></span> <span data-ttu-id="73915-233">Insira a seguinte URL para testar se a API REST está funcionando: `https://localhost:<ssl port number>/api/analyzeunicode?value=test`.</span><span class="sxs-lookup"><span data-stu-id="73915-233">Enter the following URL to test that the REST API is working: `https://localhost:<ssl port number>/api/analyzeunicode?value=test`.</span></span> <span data-ttu-id="73915-234">Você pode reutilizar o número da porta na URL no navegador que o Visual Studio iniciou.</span><span class="sxs-lookup"><span data-stu-id="73915-234">You can reuse the port number from the URL in the browser that Visual Studio launched.</span></span> <span data-ttu-id="73915-235">Você deverá ver uma cadeia de caracteres retornada com valores Unicode para cada caractere.</span><span class="sxs-lookup"><span data-stu-id="73915-235">You should see a string returned with Unicode values for each character.</span></span>

## <a name="create-the-office-add-in"></a><span data-ttu-id="73915-236">Criar o suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="73915-236">Create the Office add-in</span></span>

<span data-ttu-id="73915-237">Quando você cria o suplemento do Office, ele faz uma chamada para a API REST.</span><span class="sxs-lookup"><span data-stu-id="73915-237">When you create the Office add-in, it will make a call to the REST API.</span></span> <span data-ttu-id="73915-238">Mas, primeiro, você precisa obter o número da porta do servidor da API REST e salvá-lo para mais tarde.</span><span class="sxs-lookup"><span data-stu-id="73915-238">But first, you need to get the port number of the REST API server and save it for later.</span></span>

### <a name="save-the-ssl-port-number"></a><span data-ttu-id="73915-239">Salve o número da porta SSL</span><span class="sxs-lookup"><span data-stu-id="73915-239">Save the SSL port number</span></span>

1. <span data-ttu-id="73915-240">Caso ainda não o tenha feito, inicie o Visual Studio 2019 e abra a solução **\start\Cell-Analyzer.sln**.</span><span class="sxs-lookup"><span data-stu-id="73915-240">If you haven't already, start Visual Studio 2019, and open the **\start\Cell-Analyzer.sln** solution.</span></span>
2. <span data-ttu-id="73915-241">No projeto **CellAnalyzerRESTAPI**, expanda **Propriedades** e abra o arquivo **launchSettings. JSON**.</span><span class="sxs-lookup"><span data-stu-id="73915-241">In the **CellAnalyzerRESTAPI** project, expand **Properties**, and open the **launchSettings.json** file.</span></span>
3. <span data-ttu-id="73915-242">Localize a linha de código com o valor **sslPort**, copie o número da porta e salve-o em algum lugar.</span><span class="sxs-lookup"><span data-stu-id="73915-242">Find the line of code with the **sslPort** value, copy the port number, and save it somewhere.</span></span>

### <a name="add-the-office-add-in-project"></a><span data-ttu-id="73915-243">Adicione o projeto de suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="73915-243">Add the Office add-in project</span></span>

<span data-ttu-id="73915-244">Para simplificar, mantenha todo o código em uma solução.</span><span class="sxs-lookup"><span data-stu-id="73915-244">To keep things simple, keep all the code in one solution.</span></span> <span data-ttu-id="73915-245">Adicione o projeto do suplemento do Office à solução existente do Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="73915-245">Add the Office add-in project to the existing Visual Studio solution.</span></span> <span data-ttu-id="73915-246">No entanto, se você estiver familiarizado com o [Gerador Yeoman de Suplementos do Office](https://github.com/OfficeDev/generator-office) e do Código do Visual Studio, também poderá executar `yo office` para criar o projeto.</span><span class="sxs-lookup"><span data-stu-id="73915-246">However, if you are familiar with the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) and Visual Studio Code you can also run `yo office` to build the project.</span></span> <span data-ttu-id="73915-247">As etapas são muito semelhantes.</span><span class="sxs-lookup"><span data-stu-id="73915-247">The steps are very similar.</span></span>

1. <span data-ttu-id="73915-248">No **Gerenciador de soluções**, clique com o botão direito do mouse na solução **Cell-Analyzer** e escolha **Adicionar > Novo projeto**.</span><span class="sxs-lookup"><span data-stu-id="73915-248">In **Solution Explorer**, right-click the **Cell-Analyzer** solution, and choose **Add > New Project**.</span></span>
2. <span data-ttu-id="73915-249">Na **caixa de diálogo Adicionar um novo projeto**, clique em **Suplemento do Web Add-in** e escolha **Próximo**.</span><span class="sxs-lookup"><span data-stu-id="73915-249">In the **Add a new project dialog**, choose **Excel Web Add-in**, and choose **Next**.</span></span>
3. <span data-ttu-id="73915-250">Na caixa de diálogo **Configure seu novo projeto**, defina os seguintes campos:</span><span class="sxs-lookup"><span data-stu-id="73915-250">In the **Configure your new project** dialog, set the following fields:</span></span>
    - <span data-ttu-id="73915-251">Defina o **nome do projeto** como **CellAnalyzerOfficeAddin**.</span><span class="sxs-lookup"><span data-stu-id="73915-251">Set the **Project name** to **CellAnalyzerOfficeAddin**.</span></span>
    - <span data-ttu-id="73915-252">Deixe o **Local** com o valor padrão.</span><span class="sxs-lookup"><span data-stu-id="73915-252">Leave the **Location** at it's default value.</span></span>
    - <span data-ttu-id="73915-253">Defina a **estrutura** como **4.7.2** ou superior.</span><span class="sxs-lookup"><span data-stu-id="73915-253">Set the **Framework** to **4.7.2** or later.</span></span>
4. <span data-ttu-id="73915-254">Escolha **Criar**.</span><span class="sxs-lookup"><span data-stu-id="73915-254">Choose **Create**.</span></span>
5. <span data-ttu-id="73915-255">Na caixa de diálogo **Escolha o tipo de suplemento**, selecione **Adicionar novas funcionalidades ao Excel** e escolha **Concluir**.</span><span class="sxs-lookup"><span data-stu-id="73915-255">In the **Choose the add-in type** dialog, select **Add new functionalities to Excel**, and choose **Finish**.</span></span>

<span data-ttu-id="73915-256">Dois projetos serão criados:</span><span class="sxs-lookup"><span data-stu-id="73915-256">Two projects will be created:</span></span>

- <span data-ttu-id="73915-257">**CellAnalyzerOfficeAddin** - este projeto configura os arquivos XML de manifesto que descrevem o suplemento, para que o Office possa carregá-lo corretamente.</span><span class="sxs-lookup"><span data-stu-id="73915-257">**CellAnalyzerOfficeAddin** - This project configures the manifest XML files that describes the add-in so Office can load it correctly.</span></span> <span data-ttu-id="73915-258">Ele contém o ID, nome, descrição e outras informações sobre o suplemento.</span><span class="sxs-lookup"><span data-stu-id="73915-258">It contains the ID, name, description, and other information about the add-in.</span></span>
- <span data-ttu-id="73915-259">**CellAnalyzerOfficeAddinWeb** - este projeto contém recursos da Web para seu suplemento, como HTML, CSS e scripts.</span><span class="sxs-lookup"><span data-stu-id="73915-259">**CellAnalyzerOfficeAddinWeb** - This project contains web resources for your add-in, such as HTML, CSS, and scripts.</span></span> <span data-ttu-id="73915-260">Ele também configura uma instância do IIS Express para hospedar seu suplemento como um aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="73915-260">It also configures an IIS Express instance to host your add-in as a web application.</span></span>

### <a name="add-ui-and-functionality-to-the-office-add-in"></a><span data-ttu-id="73915-261">Adicionar interface de usuário e funcionalidade ao suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="73915-261">Add UI and functionality to the Office add-in</span></span>

1. <span data-ttu-id="73915-262">No **Gerenciador de soluções**, expanda o projeto **CellAnalyzerOfficeAddinWeb**.</span><span class="sxs-lookup"><span data-stu-id="73915-262">In **Solution Explorer**, expand the **CellAnalyzerOfficeAddinWeb** project.</span></span>
2. <span data-ttu-id="73915-263">Abra o arquivo **Home.HTML** e substitua o conteúdo de `<body>` pela seguinte HTML.</span><span class="sxs-lookup"><span data-stu-id="73915-263">Open the **Home.html** file, and replace the `<body>` contents with the following HTML.</span></span>

    ```html
    <button id="btnShowUnicode" onclick="showUnicode()">Show Unicode</button>
    <p>Result:</p>
    <div id="txtResult"></div>
    ```

3. <span data-ttu-id="73915-264">Abra o arquivo **Home.js** e substitua todo o conteúdo pelo seguinte código.</span><span class="sxs-lookup"><span data-stu-id="73915-264">Open the **Home.js** file, and replace the entire contents with the following code.</span></span>

    ```js
    (function () {
      "use strict";
      // The initialize function must be run each time a new page is loaded.
      Office.initialize = function (reason) {
        $(document).ready(function () {
        });
      };
    })();

    function showUnicode() {
      Excel.run(function (ctx) {
        const range = ctx.workbook.getSelectedRange();
        range.load("values");
        return ctx.sync(range).then(function (range) {
          const url = "https://localhost:<ssl port number>/api/analyzeunicode?value=" + range.values[0][0];
          $.ajax({
            type: "GET",
            url: url,
            success: function (data) {
              let htmlData = data.replace(/\r\n/g, '<br>');
              $("#txtResult").html(htmlData);
            },
            error: function (data) {
                $("#txtResult").html("error occurred in ajax call.");
            }
          });
        });
      });
    }
    ```

4. <span data-ttu-id="73915-265">No código anterior, digite o número **sslPort** que você salvou anteriormente pelo arquivo **. JSON**.</span><span class="sxs-lookup"><span data-stu-id="73915-265">In the previous code, enter the **sslPort** number you saved previously from the **launchSettings.json** file.</span></span>

<span data-ttu-id="73915-266">No código anterior, a cadeia de caracteres retornada será processada para substituir alimentações de linha de retorno de carro por marcas `<br>` HTML.</span><span class="sxs-lookup"><span data-stu-id="73915-266">In the previous code the returned string will be processed to replace carriage return line feeds with `<br>` HTML tags.</span></span> <span data-ttu-id="73915-267">Algumas vezes, você pode encontrar situações em que um valor de retorno que funcione perfeitamente para o .NET precisará ser ajustado no suplemento do Office para trabalhar conforme o esperado no suplemento VSTO .</span><span class="sxs-lookup"><span data-stu-id="73915-267">You may occasionally run into situations where a return value that works perfectly fine for .NET in the VSTO Add-in will need to be adjusted on the Office add-in side to work as expected.</span></span> <span data-ttu-id="73915-268">Nesse caso, a API REST e a biblioteca de classes compartilhadas só se preocupam em retornar a cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="73915-268">In this case the REST API and shared class library are only concerned with returning the string.</span></span> <span data-ttu-id="73915-269">O método `showUnicode()` é responsável pela formatação de valores retornados corretamente para a apresentação.</span><span class="sxs-lookup"><span data-stu-id="73915-269">The `showUnicode()` method is responsible for formatting return values correctly for presentation.</span></span>

### <a name="allow-cors-from-the-office-add-in"></a><span data-ttu-id="73915-270">Permitir CORS no suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="73915-270">Allow CORS from the Office add-in</span></span>

<span data-ttu-id="73915-271">A biblioteca do Office. js exige o CORS nas chamadas de saída, como a realizada na chamada `ajax` para o servidor de API REST.</span><span class="sxs-lookup"><span data-stu-id="73915-271">The Office.js library requires CORS on outgoing calls, such as the one made from the `ajax` call to the REST API server.</span></span> <span data-ttu-id="73915-272">Use as etapas a seguir para permitir chamadas do suplemento do Office para a API REST.</span><span class="sxs-lookup"><span data-stu-id="73915-272">Use the following steps to allow calls from the Office add-in to the REST API.</span></span>

1. <span data-ttu-id="73915-273">No **Gerenciador de soluções**, selecione o projeto **CellAnalyzerOfficeAddinWeb**.</span><span class="sxs-lookup"><span data-stu-id="73915-273">In **Solution Explorer**, select the **CellAnalyzerOfficeAddinWeb** project.</span></span>
2. <span data-ttu-id="73915-274">No menu **Exibir**, escolha **Janela Propriedades** (se a janela ainda não estiver sendo exibida).</span><span class="sxs-lookup"><span data-stu-id="73915-274">From the **View** menu, choose **Properties Window** (if the window is not already displayed).</span></span>
3. <span data-ttu-id="73915-275">Na janela Propriedades, copie o valor da URL **SSL** e salve-a em outro local.</span><span class="sxs-lookup"><span data-stu-id="73915-275">In the properties window, copy the value of the **SSL URL**, and save it somewhere.</span></span> <span data-ttu-id="73915-276">Esta é a URL necessária para permitir o CORS.</span><span class="sxs-lookup"><span data-stu-id="73915-276">This is the URL that you need to allow through CORS.</span></span>
4. <span data-ttu-id="73915-277">No projeto **CellAnalyzerRESTAPI**, abra o arquivo **Startup.cs**.</span><span class="sxs-lookup"><span data-stu-id="73915-277">In the **CellAnalyzerRESTAPI** project, open the **Startup.cs** file.</span></span>
5. <span data-ttu-id="73915-278">Na parte superior do método, adicione o seguinte código `ConfigureServices`.</span><span class="sxs-lookup"><span data-stu-id="73915-278">Add the following code to the top of the `ConfigureServices` method.</span></span> <span data-ttu-id="73915-279">Substitua a URL SSL que você copiou anteriormente para a chamada `builder.WithOrigins`.</span><span class="sxs-lookup"><span data-stu-id="73915-279">Be sure to substitute the URL SSL you copied previously for the `builder.WithOrigins` call.</span></span>

    ```csharp
    services.AddCors(options =>
    {
      options.AddPolicy(MyAllowSpecificOrigins,
      builder =>
      {
        builder.WithOrigins("<your URL SSL>")
        .AllowAnyMethod()
        .AllowAnyHeader();
      });
    });
    ```

    > [!NOTE]
    > <span data-ttu-id="73915-280">Mantenha o final `/` da URL ao usá-lo no método `builder.WithOrigins`.</span><span class="sxs-lookup"><span data-stu-id="73915-280">Leave the trailing `/` from the end of the URL when you use it in the `builder.WithOrigins` method.</span></span> <span data-ttu-id="73915-281">Por exemplo, ele deve parecer semelhante a `https://localhost:44000`.</span><span class="sxs-lookup"><span data-stu-id="73915-281">For example, it should appear similar to `https://localhost:44000`.</span></span> <span data-ttu-id="73915-282">Caso contrário, você receberá um erro CORS em tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="73915-282">Otherwise you will get a CORS error at runtime.</span></span>

6. <span data-ttu-id="73915-283">Adicione o campo a seguir à classe `Startup`:</span><span class="sxs-lookup"><span data-stu-id="73915-283">Add the following field to the `Startup` class:</span></span>

    ```csharp
    readonly string MyAllowSpecificOrigins = "_myAllowSpecificOrigins";
    ```

7. <span data-ttu-id="73915-284">Adicione o seguinte código ao método `configure` logo antes da linha de código para `app.UseEndpoints`.</span><span class="sxs-lookup"><span data-stu-id="73915-284">Add the following code to the `configure` method just before the line of code for `app.UseEndpoints`.</span></span>

    ```csharp
    app.UseCors(MyAllowSpecificOrigins);
    ```

<span data-ttu-id="73915-285">Quando terminar, sua classe `Startup` deve ser semelhante ao seguinte código (sua URL de localhost pode ser diferente):</span><span class="sxs-lookup"><span data-stu-id="73915-285">When done, your `Startup` class should look similar to the following code (your localhost URL may be different):</span></span>

```csharp
public class Startup
{
  public Startup(IConfiguration configuration)
    {
      Configuration = configuration;
    }

    readonly string MyAllowSpecificOrigins = "_myAllowSpecificOrigins";

    public IConfiguration Configuration { get; }

    // NOTE: The following code configures CORS for the localhost:44397 port.
    // This is for development purposes. In production code you should update this to 
    // use the appropriate allowed domains.
    public void ConfigureServices(IServiceCollection services)
    {
        services.AddCors(options =>
        {
            options.AddPolicy(MyAllowSpecificOrigins,
            builder =>
            {
                builder.WithOrigins("https://localhost:44397")
                .AllowAnyMethod()
                .AllowAnyHeader();
            });
        });
        services.AddControllers();
    }

    // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
    public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
    {
        if (env.IsDevelopment())
        {
            app.UseDeveloperExceptionPage();
        }

        app.UseHttpsRedirection();

        app.UseRouting();

        app.UseAuthorization();

        app.UseCors(MyAllowSpecificOrigins);

        app.UseEndpoints(endpoints =>
        {
            endpoints.MapControllers();
        });
    }
}
```

### <a name="run-the-add-in"></a><span data-ttu-id="73915-286">Execute o suplemento</span><span class="sxs-lookup"><span data-stu-id="73915-286">Run the add-in</span></span>

1. <span data-ttu-id="73915-287">No **Explorador de Soluções**, clique com o botão direito do mouse no nó superior **Solução 'Cell-Analyzer'** e escolha **Definir Projetos de Inicialização**.</span><span class="sxs-lookup"><span data-stu-id="73915-287">In **Solution Explorer**, right-click the top node **Solution 'Cell-Analyzer'**, and choose **Set Startup Projects**.</span></span>
2. <span data-ttu-id="73915-288">Na caixa de diálogo **Páginas de propriedades da solução 'Cell-Analyzer'**, selecione **Vários projetos de inicialização**.</span><span class="sxs-lookup"><span data-stu-id="73915-288">In the **Solution 'Cell-Analyzer' Property Pages** dialog, select **Multiple startup projects**.</span></span>
3. <span data-ttu-id="73915-289">Defina a propriedade **Action** como **Iniciar** para cada um dos seguintes projetos.</span><span class="sxs-lookup"><span data-stu-id="73915-289">Set the **Action** property to **Start** for each of the following projects.</span></span>

    - <span data-ttu-id="73915-290">CellAnalyzerRESTAPI</span><span class="sxs-lookup"><span data-stu-id="73915-290">CellAnalyzerRESTAPI</span></span>
    - <span data-ttu-id="73915-291">CellAnalyzerOfficeAddin</span><span class="sxs-lookup"><span data-stu-id="73915-291">CellAnalyzerOfficeAddin</span></span>
    - <span data-ttu-id="73915-292">CellAnalyzerOfficeAddinWeb</span><span class="sxs-lookup"><span data-stu-id="73915-292">CellAnalyzerOfficeAddinWeb</span></span>

4. <span data-ttu-id="73915-293">Escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="73915-293">Choose **OK**.</span></span>
5. <span data-ttu-id="73915-294">No menu **Depurar**, selecione **Iniciar Depuração**.</span><span class="sxs-lookup"><span data-stu-id="73915-294">From the **Debug** menu, choose **Start Debugging**.</span></span>

<span data-ttu-id="73915-295">O Excel será executado e fará o carregamento lateral do suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="73915-295">Excel will run and sideload the Office add-in.</span></span> <span data-ttu-id="73915-296">Você pode testar se o serviço de API do localhost REST está funcionando corretamente, inserindo um valor de texto em uma célula e escolhendo o botão **Mostrar Unicode** no suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="73915-296">You can test that the localhost REST API service is working correctly by entering a text value into a cell, and choosing the **Show Unicode** button in the Office add-in.</span></span> <span data-ttu-id="73915-297">Ele deve chamar a API REST e exibir os valores Unicode para os caracteres de texto.</span><span class="sxs-lookup"><span data-stu-id="73915-297">It should call the REST API and display the unicode values for the text characters.</span></span>

## <a name="publish-to-an-azure-app-service"></a><span data-ttu-id="73915-298">Publicar em um serviço de aplicativo do Azure</span><span class="sxs-lookup"><span data-stu-id="73915-298">Publish to an Azure App Service</span></span>

<span data-ttu-id="73915-299">Eventualmente, você deseja publicar o projeto da API REST na nuvem.</span><span class="sxs-lookup"><span data-stu-id="73915-299">You eventually want to publish the REST API project to the cloud.</span></span> <span data-ttu-id="73915-300">Nas etapas a seguir, você verá como publicar o projeto **CellAnalyzerRESTAPI** em um serviço de aplicativo do Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="73915-300">In the following steps you'll see how to publish the **CellAnalyzerRESTAPI** project to a Microsoft Azure App Service.</span></span> <span data-ttu-id="73915-301">Confira os[pré-requisitos](#prerequisites) para saber mais sobre como obter uma conta do Azure.</span><span class="sxs-lookup"><span data-stu-id="73915-301">See [Prerequisites](#prerequisites) for information on how to get an Azure account.</span></span>

1. <span data-ttu-id="73915-302">No **Gerenciador de soluções**, clique com o botão direito do mouse no projeto **CellAnalyzerRESTAPI** e escolha **Publicar**.</span><span class="sxs-lookup"><span data-stu-id="73915-302">In **Solution Explorer**, right-click the **CellAnalyzerRESTAPI** project, and choose **Publish**.</span></span>
2. <span data-ttu-id="73915-303">Na caixa de diálogo **Escolha um destino de publicação**, selecione **Criar Novo** e escolha **Criar Perfil**.</span><span class="sxs-lookup"><span data-stu-id="73915-303">In the **Pick a publish target** dialog, select **Create New**, and choose **Create Profile**.</span></span>
3. <span data-ttu-id="73915-304">Na caixa de diálogo **Serviço de Aplicativo**, selecione a conta correta, caso ainda não esteja selecionada.</span><span class="sxs-lookup"><span data-stu-id="73915-304">In the **App Service** dialog, select the correct account, if it is not already selected.</span></span>
4. <span data-ttu-id="73915-305">Os campos para a caixa de diálogo **Serviço de Aplicativo** serão definidos como padrões para a sua conta.</span><span class="sxs-lookup"><span data-stu-id="73915-305">The fields for the **App Service** dialog will be set to defaults for your account.</span></span> <span data-ttu-id="73915-306">Geralmente, os padrões funcionam corretamente, mas você pode alterá-los caso prefira configurações diferentes.</span><span class="sxs-lookup"><span data-stu-id="73915-306">Generally the defaults work fine, but you can change them if you prefer different settings.</span></span>
5. <span data-ttu-id="73915-307">Na caixa de diálogo **Serviço de Aplicativo**, escolha **Criar**.</span><span class="sxs-lookup"><span data-stu-id="73915-307">In the **App Service** dialog, choose **Create**.</span></span>
6. <span data-ttu-id="73915-308">O novo perfil será exibido em uma página de **Publicação**.</span><span class="sxs-lookup"><span data-stu-id="73915-308">The new profile will be displayed in a **Publish** page.</span></span> <span data-ttu-id="73915-309">Escolha **Publicar** para criar e implantar o código no serviço de aplicativo.</span><span class="sxs-lookup"><span data-stu-id="73915-309">Choose **Publish** to build and deploy the code to the App Service.</span></span>

<span data-ttu-id="73915-310">Agora você pode testar o serviço.</span><span class="sxs-lookup"><span data-stu-id="73915-310">You can now test the service.</span></span> <span data-ttu-id="73915-311">Abra um navegador e insira uma URL que vai diretamente para o novo serviço.</span><span class="sxs-lookup"><span data-stu-id="73915-311">Open a browser and enter a URL that goes directly to the new service.</span></span> <span data-ttu-id="73915-312">Por exemplo, use `https://<myappservice>.azurewebsites.net/api/analyzeunicode?value=test` onde *myappservice* é o nome exclusivo que você criou para o novo serviço de aplicativo.</span><span class="sxs-lookup"><span data-stu-id="73915-312">For example, use `https://<myappservice>.azurewebsites.net/api/analyzeunicode?value=test` where *myappservice* is the unique name you created for the new App Service.</span></span>

### <a name="use-the-azure-app-service-from-the-office-add-in"></a><span data-ttu-id="73915-313">Usar o serviço de aplicativo Azure do suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="73915-313">Use the Azure App Service from the Office add-in</span></span>

<span data-ttu-id="73915-314">A etapa final é atualizar o código no suplemento do Office para usar o serviço do aplicativo Azure, em vez de localhost.</span><span class="sxs-lookup"><span data-stu-id="73915-314">The final step is to update the code in the Office add-in to use the Azure App Service instead of localhost.</span></span>

1. <span data-ttu-id="73915-315">No **Gerenciador de soluções**, expanda o projeto **CellAnalyzerOfficeAddinWeb** e abra o arquivo **Home. js**.</span><span class="sxs-lookup"><span data-stu-id="73915-315">In **Solution Explorer**, expand the **CellAnalyzerOfficeAddinWeb** project, and open the **Home.js** file.</span></span>
1. <span data-ttu-id="73915-316">Altere a constante `url` para usar a URL do serviço do aplicativo Azure, como mostra a linha de código a seguir.</span><span class="sxs-lookup"><span data-stu-id="73915-316">Change the `url` constant to use the URL for your Azure App Service as shown in the following line of code.</span></span> <span data-ttu-id="73915-317">Substitua `<myappservice>` pelo nome exclusivo que você criou para o novo serviço de aplicativo.</span><span class="sxs-lookup"><span data-stu-id="73915-317">Replace `<myappservice>` with the unique name you created for the new App Service.</span></span>

    ```JavaScript
    const url = "https://<myappservice>.azurewebsites.net/api/analyzeunicode?value=" + range.values[0][0];
    ```

1. <span data-ttu-id="73915-318">No **Explorador de Soluções**, clique com o botão direito do mouse no nó superior **Solução 'Cell-Analyzer'** e escolha **Definir Projetos de Inicialização**.</span><span class="sxs-lookup"><span data-stu-id="73915-318">In **Solution Explorer**, right-click the top node **Solution 'Cell-Analyzer'**, and choose **Set Startup Projects**.</span></span>
1. <span data-ttu-id="73915-319">Na caixa de diálogo **Páginas de propriedades da solução 'Cell-Analyzer'**, selecione **Vários projetos de inicialização**.</span><span class="sxs-lookup"><span data-stu-id="73915-319">In the **Solution 'Cell-Analyzer' Property Pages** dialog, select **Multiple startup projects**.</span></span>
1. <span data-ttu-id="73915-320">Habilite a ação **Iniciar** para cada um dos seguintes projetos:</span><span class="sxs-lookup"><span data-stu-id="73915-320">Enable the **Start** action for each of the following projects:</span></span>
    - <span data-ttu-id="73915-321">CellAnalyzerOfficeAddinWeb</span><span class="sxs-lookup"><span data-stu-id="73915-321">CellAnalyzerOfficeAddinWeb</span></span>
    - <span data-ttu-id="73915-322">CellAnalyzerOfficeAddin</span><span class="sxs-lookup"><span data-stu-id="73915-322">CellAnalyzerOfficeAddin</span></span>
1. <span data-ttu-id="73915-323">Escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="73915-323">Choose **OK**.</span></span>
1. <span data-ttu-id="73915-324">No menu **Depurar**, selecione **Iniciar Depuração**.</span><span class="sxs-lookup"><span data-stu-id="73915-324">From the **Debug** menu, choose **Start Debugging**.</span></span>

<span data-ttu-id="73915-325">O Excel será executado e fará o carregamento lateral do suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="73915-325">Excel will run and sideload the Office add-in.</span></span> <span data-ttu-id="73915-326">Para testar se o serviço de aplicativo está funcionando corretamente, insira um valor de texto em uma célula e escolha **Mostrar Unicode** no suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="73915-326">To test that the App Service is working correctly, enter a text value into a cell, and choose **Show Unicode** in the Office add-in.</span></span> <span data-ttu-id="73915-327">Ele deve chamar o serviço e exibir os valores Unicode para os caracteres de texto.</span><span class="sxs-lookup"><span data-stu-id="73915-327">It should call the service and display the unicode values for the text characters.</span></span>

## <a name="conclusion"></a><span data-ttu-id="73915-328">Conclusão</span><span class="sxs-lookup"><span data-stu-id="73915-328">Conclusion</span></span>

<span data-ttu-id="73915-329">Neste tutorial você aprendeu a criar um suplemento do Office que usa um código compartilhado com o suplemento VSTO original.</span><span class="sxs-lookup"><span data-stu-id="73915-329">In this tutorial you learned how to create an Office add-in that uses shared code with the original VSTO add-in.</span></span> <span data-ttu-id="73915-330">Você aprendeu como manter o código VSTO do Office no Windows e um suplemento do Office para o Office em outras plataformas.</span><span class="sxs-lookup"><span data-stu-id="73915-330">You learned how to maintain both VSTO code for Office on Windows, and an Office add-in for Office on other platforms.</span></span> <span data-ttu-id="73915-331">Você refatorou o código C # do VSTO em uma biblioteca compartilhada e o implantou em um Serviço de Aplicativo do Azure.</span><span class="sxs-lookup"><span data-stu-id="73915-331">You refactored VSTO C# code into a shared library and deployed it to an Azure App Service.</span></span> <span data-ttu-id="73915-332">Você criou um suplemento do Office que usa a biblioteca compartilhadas para que não seja necessário regravar o código em JavaScript.</span><span class="sxs-lookup"><span data-stu-id="73915-332">You created an Office add-in that uses the shared library so that you don't have to rewrite the code in JavaScript.</span></span>
