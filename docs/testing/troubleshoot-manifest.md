---
title: Validar e solucionar problemas com seu manifesto
description: Use estes métodos para validar o manifesto de Suplementos do Office.
ms.date: 12/04/2017
ms.openlocfilehash: 51d644f7cfb7fbad5c9b66be41dc57015202b9be
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/19/2018
ms.locfileid: "25639984"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="71ffa-103">Validar e solucionar problemas com seu manifesto</span><span class="sxs-lookup"><span data-stu-id="71ffa-103">Validate and troubleshoot issues with your manifest</span></span>

<span data-ttu-id="71ffa-104">Use estes métodos para validar e solucionar problemas no manifesto dos Suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="71ffa-104">Use these methods to validate and troubleshoot issues in your Office Add-ins manifest:</span></span> 

- [<span data-ttu-id="71ffa-105">Validar o manifesto com o Validador de Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="71ffa-105">Validate your manifest with the Office Add-in Validator</span></span>](#validate-your-manifest-with-the-office-add-in-validator)   
- [<span data-ttu-id="71ffa-106">Validar o manifesto com base no esquema XML</span><span class="sxs-lookup"><span data-stu-id="71ffa-106">Validate your manifest against the XML schema</span></span>](#validate-your-manifest-against-the-xml-schema)
- [<span data-ttu-id="71ffa-107">Usar o log de tempo de execução para depurar o manifesto do suplemento</span><span class="sxs-lookup"><span data-stu-id="71ffa-107">Use runtime logging to debug your add-in manifest</span></span>](#use-runtime-logging-to-debug-your-add-in-manifest)


## <a name="validate-your-manifest-with-the-office-add-in-validator"></a><span data-ttu-id="71ffa-108">Validar o manifesto com o Validador de Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="71ffa-108">Validate your manifest with the Office Add-in Validator</span></span>

<span data-ttu-id="71ffa-109">Para se assegurar de que o arquivo de manifesto que descreve o Suplemento do Office está correto e completo, valide-o com o [Validador de Suplemento do Office](https://github.com/OfficeDev/office-addin-validator).</span><span class="sxs-lookup"><span data-stu-id="71ffa-109">To help ensure that the manifest file that describes your Office Add-in is correct and complete, validate it against the [Office Add-in Validator](https://github.com/OfficeDev/office-addin-validator).</span></span>

### <a name="to-use-the-office-add-in-validator-to-validate-your-manifest"></a><span data-ttu-id="71ffa-110">Para usar o Validador de Suplemento do Office para validar o manifesto</span><span class="sxs-lookup"><span data-stu-id="71ffa-110">To use the Office Add-in Validator to validate your manifest</span></span>

1. <span data-ttu-id="71ffa-111">Instale o [Node.js](https://nodejs.org/download/).</span><span class="sxs-lookup"><span data-stu-id="71ffa-111">Install [Node.js](https://nodejs.org/download/).</span></span> 

2. <span data-ttu-id="71ffa-112">Abra um prompt de comando/terminal como administrador e instale o Validador de Suplemento do Office e suas dependências globalmente usando o seguinte comando:</span><span class="sxs-lookup"><span data-stu-id="71ffa-112">Open a command prompt / terminal as an administrator, and install the Office Add-in Validator and its dependencies globally by using the following command:</span></span>

    ```bash
    npm install -g office-addin-validator
    ```
    
    > [!NOTE]
    > <span data-ttu-id="71ffa-113">se já instalou o Yo Office, atualize-o para a última versão e o validador será instalado como uma dependência.</span><span class="sxs-lookup"><span data-stu-id="71ffa-113">If you already have Yo Office installed, upgrade to the latest version, and the validator will be installed as a dependency.</span></span>

3. <span data-ttu-id="71ffa-p101">Para validar o manifesto, execute o seguinte comando. Substitua MANIFEST.XML pelo caminho para o arquivo XML do manifesto.</span><span class="sxs-lookup"><span data-stu-id="71ffa-p101">Run the following command to validate your manifest. Replace MANIFEST.XML with the path to the manifest XML file.</span></span>

    ```bash
    validate-office-addin MANIFEST.XML
    ```

## <a name="validate-your-manifest-against-the-xml-schema"></a><span data-ttu-id="71ffa-116">Validar o manifesto com base no esquema XML</span><span class="sxs-lookup"><span data-stu-id="71ffa-116">Validate your manifest against the XML schema</span></span>

<span data-ttu-id="71ffa-p102">Para ajudar a garantir que o arquivo de manifesto segue o esquema correto, incluindo quaisquer namespaces para os elementos que você está usando. Se você copiou elementos de outros manifestos de amostra, verifique se também **incluem os namespaces apropriados**. Você pode validar um manifesto com base nos arquivos da [Definição de Esquema XML  (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas). Pode usar uma ferramenta de validação de esquema XML para executar essa validação.</span><span class="sxs-lookup"><span data-stu-id="71ffa-p102">To help ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using. If you copied elements from other sample manifests double check you also **include the appropiate namespaces**. You can validate a manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) files. You can use an XML schema validation tool to perform this validation.</span></span> 



### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a><span data-ttu-id="71ffa-121">Usar uma ferramenta de validação de esquema XML da linha de comando para validar o manifesto</span><span class="sxs-lookup"><span data-stu-id="71ffa-121">To use a command-line XML schema validation tool to validate your manifest</span></span>

1.  <span data-ttu-id="71ffa-122">Instale o [tar](https://www.gnu.org/software/tar/) e o [libxml](http://xmlsoft.org/FAQ.html), caso ainda não tenha instalado.</span><span class="sxs-lookup"><span data-stu-id="71ffa-122">Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.</span></span>

2.  <span data-ttu-id="71ffa-p103">Execute o comando a seguir. Substitua `XSD_FILE` pelo caminho para o arquivo XSD do manifesto e `XML_FILE` pelo caminho para o arquivo XML do manifesto.</span><span class="sxs-lookup"><span data-stu-id="71ffa-p103">Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.</span></span>
    
    ```bash
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="use-runtime-logging-to-debug-your-add-in"></a><span data-ttu-id="71ffa-125">Use o log de tempo de execução para depurar o suplemento</span><span class="sxs-lookup"><span data-stu-id="71ffa-125">Use runtime logging to debug your add-in manifest</span></span> 

<span data-ttu-id="71ffa-p104">Você pode usar o log de tempo de execução para depurar o manifesto do suplemento, bem como vários erros de instalação. Esse recurso pode ajudá-lo a identificar e corrigir problemas com o manifesto que não são detectados pela validação do esquema XSD, como uma incompatibilidade entre os IDs do recurso. O log de tempo de execução é particularmente útil para depurar suplementos que implementam comandos de suplementos e funções personalizadas do Excel.</span><span class="sxs-lookup"><span data-stu-id="71ffa-p104">You can use runtime logging to debug your add-in's manifest as well as several installation errors. This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs. Runtime logging is particularly  useful for debugging add-ins that implement add-in commands and Excel custom functions.</span></span>   

> [!NOTE]
> <span data-ttu-id="71ffa-129">O recurso de log de tempo de execução está disponível atualmente para o Office 2016 desktop.</span><span class="sxs-lookup"><span data-stu-id="71ffa-129">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

### <a name="to-turn-on-runtime-logging"></a><span data-ttu-id="71ffa-130">Para ativar o log de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="71ffa-130">To turn on runtime logging</span></span>

> [!IMPORTANT]
> <span data-ttu-id="71ffa-p105">O log do tempo de execução afeta o desempenho. Ative-o somente quando precisar depurar problemas com o manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="71ffa-p105">Runtime Logging affects performance. Turn it on only when you need to debug issues with your add-in manifest.</span></span>

<span data-ttu-id="71ffa-133">Para ativar o log de tempo de execução:</span><span class="sxs-lookup"><span data-stu-id="71ffa-133">To turn on runtime logging:</span></span>

1. <span data-ttu-id="71ffa-134">Verifique se você está executando o Office 2016 desktop, build **16.0.7019** ou posterior.</span><span class="sxs-lookup"><span data-stu-id="71ffa-134">Make sure that you are running Office 2016 desktop build **16.0.7019** or later.</span></span> 

2. <span data-ttu-id="71ffa-135">Adicione a chave do registro `RuntimeLogging` em  `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\`.</span><span class="sxs-lookup"><span data-stu-id="71ffa-135">Add the `RuntimeLogging` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\`.</span></span> 

3. <span data-ttu-id="71ffa-p106">Defina o valor padrão da chave para o caminho completo do arquivo onde você deseja que o log seja gravado. Para obter um exemplo, veja [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span><span class="sxs-lookup"><span data-stu-id="71ffa-p106">Set the default value of the key to the full path of the file where you want the log to be written. For an example, see [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="71ffa-138">O diretório no qual o arquivo de log será gravado já deve existir e você deve ter permissões de gravação.</span><span class="sxs-lookup"><span data-stu-id="71ffa-138">The directory in which the log file will be written must already exist, and you must have write permissions to it.</span></span> 
 
<span data-ttu-id="71ffa-p107">A imagem a seguir mostra como deve ficar o registro. Para desativar o recurso, remova a chave  `RuntimeLogging` do registro.</span><span class="sxs-lookup"><span data-stu-id="71ffa-p107">The following image shows what the registry should look like. To turn the feature off, remove the `RuntimeLogging` key from the registry.</span></span> 

![Captura de tela do editor de registro com uma chave de registro RuntimeLogging](http://i.imgur.com/Sa9TyI6.png)


### <a name="to-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="71ffa-142">Para solucionar problemas com o manifesto</span><span class="sxs-lookup"><span data-stu-id="71ffa-142">To troubleshoot issues with your manifest</span></span>

<span data-ttu-id="71ffa-143">Para usar o log de tempo de execução para solucionar problemas ao carregar um suplemento:</span><span class="sxs-lookup"><span data-stu-id="71ffa-143">To use runtime logging to troubleshoot issues loading an add-in:</span></span>
 
1. <span data-ttu-id="71ffa-144">[Realize o sideload do suplemento](sideload-office-add-ins-for-testing.md) para testes.</span><span class="sxs-lookup"><span data-stu-id="71ffa-144">[Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="71ffa-145">Recomendamos realizar o sideload apenas do suplemento que você está testando para diminuir a quantidade de mensagens no arquivo de log.</span><span class="sxs-lookup"><span data-stu-id="71ffa-145">We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.</span></span>

2. <span data-ttu-id="71ffa-146">Se nada acontece e se você não vê o suplemento (não aparece na caixa de diálogo de suplementos), abra o arquivo de log.</span><span class="sxs-lookup"><span data-stu-id="71ffa-146">If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.</span></span>

3. <span data-ttu-id="71ffa-p108">Procure no arquivo de log o ID de seu suplemento que foi definido no seu manifesto. No arquivo de log, esse ID está marcado como `SolutionId`.</span><span class="sxs-lookup"><span data-stu-id="71ffa-p108">Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`.</span></span> 

<span data-ttu-id="71ffa-p109">No exemplo a seguir, o arquivo de log identifica um controle que aponta para um arquivo de recurso que não existe. Neste exemplo, a correção seria corrigir o erro de digitação no manifesto ou adicionar o recurso ausente.</span><span class="sxs-lookup"><span data-stu-id="71ffa-p109">In the following example, the log file identifies a control that points to a resource file that doesn't exist. For this example, the fix would be to correct the typo in the manifest or to add the missing resource.</span></span>

![Captura de tela de um arquivo de log com uma entrada que especifica a identificação do recurso que não foi encontrado](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a><span data-ttu-id="71ffa-152">Problemas conhecidos com o log de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="71ffa-152">Known issues with runtime logging</span></span>

<span data-ttu-id="71ffa-p110">Talvez você veja mensagens no arquivo de log que são confusas ou que estão classificadas incorretamente. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="71ffa-p110">You might see messages in the log file that are confusing or that are classified incorrectly. For example:</span></span>

- <span data-ttu-id="71ffa-155">A mensagem `Medium Current host not in add-in's host list` seguida por `Unexpected Parsed manifest targeting different host` está classificada incorretamente como um erro.</span><span class="sxs-lookup"><span data-stu-id="71ffa-155">The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.</span></span>

- <span data-ttu-id="71ffa-156">Se você vir a mensagem `Unexpected Add-in is missing required manifest fields DisplayName` e ela não contiver uma SolutionId, o erro provavelmente não está relacionado ao suplemento que você está depurando.</span><span class="sxs-lookup"><span data-stu-id="71ffa-156">If you see the message `Unexpected Add-in is missing required manifest fields DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging.</span></span> 

- <span data-ttu-id="71ffa-p111">Todas as mensagens `Monitorable` indicam erros esperados do ponto de vista do sistema. Às vezes, indicam um problema com seu manifesto, como um elemento com erro de ortografia que foi ignorado, mas não causou a falha do manifesto.</span><span class="sxs-lookup"><span data-stu-id="71ffa-p111">Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail.</span></span> 

## <a name="clear-the-office-cache"></a><span data-ttu-id="71ffa-159">Limpar o cache do Office</span><span class="sxs-lookup"><span data-stu-id="71ffa-159">Clear the Office cache</span></span>

<span data-ttu-id="71ffa-160">Se as alterações feitas no manifesto, como nomes de arquivo dos ícones do botão da faixa de opções ou o texto de comandos do suplemento, não entraram em vigor, tente limpar o cache do Office no computador.</span><span class="sxs-lookup"><span data-stu-id="71ffa-160">If changes you've made in the manifest, such as file names of ribbon button icons, or text of add-in commands, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="71ffa-161">No Windows:</span><span class="sxs-lookup"><span data-stu-id="71ffa-161">For Windows:</span></span>
<span data-ttu-id="71ffa-162">Exclua o conteúdo da pasta `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="71ffa-162">Delete the content of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="71ffa-163">No Mac:</span><span class="sxs-lookup"><span data-stu-id="71ffa-163">For Mac:</span></span>
<span data-ttu-id="71ffa-164">Exclua o conteúdo da pasta `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="71ffa-164">Delete the content of the folder `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>

#### <a name="for-ios"></a><span data-ttu-id="71ffa-165">No iOS:</span><span class="sxs-lookup"><span data-stu-id="71ffa-165">For iOS:</span></span>
<span data-ttu-id="71ffa-p112">Chame o `window.location.reload(true)` do JavaScript no suplemento para forçar um recarregamento. Outra alternativa é reinstalar o Office.</span><span class="sxs-lookup"><span data-stu-id="71ffa-p112">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="71ffa-168">Confira também</span><span class="sxs-lookup"><span data-stu-id="71ffa-168">See also</span></span>

- [<span data-ttu-id="71ffa-169">Manifesto XML dos suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="71ffa-169">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="71ffa-170">Fazer sideload de Suplementos do Office para testes</span><span class="sxs-lookup"><span data-stu-id="71ffa-170">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="71ffa-171">Depurar suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="71ffa-171">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
