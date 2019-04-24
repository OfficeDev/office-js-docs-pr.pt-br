---
title: Validar e solucionar problemas com seu manifesto
description: Use estes métodos para validar o manifesto de suplementos do Office
ms.date: 11/02/2018
localization_priority: Priority
ms.openlocfilehash: 921adf6f1f398887d96031790facc1fb1425af2b
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451148"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="abef8-103">Validar e solucionar problemas com seu manifesto</span><span class="sxs-lookup"><span data-stu-id="abef8-103">Validate and troubleshoot issues with your manifest</span></span>

<span data-ttu-id="abef8-104">Use esses métodos para validar e solucionar problemas no manifesto de seu suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="abef8-104">Use these methods to validate and troubleshoot issues in your Office Add-ins manifest:</span></span> 

- [<span data-ttu-id="abef8-105">Validar o manifesto com o Validador de Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="abef8-105">Validate your manifest with the Office Add-in Validator</span></span>](#validate-your-manifest-with-the-office-add-in-validator)   
- [<span data-ttu-id="abef8-106">Validar seu manifesto em relação ao esquema XML</span><span class="sxs-lookup"><span data-stu-id="abef8-106">Validate your manifest against the XML schema</span></span>](#validate-your-manifest-against-the-xml-schema)
- [<span data-ttu-id="abef8-107">Validar o manifesto com o gerador Yeoman para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="abef8-107">Validate your manifest with the Yeoman generator for Office Add-ins</span></span>](#validate-your-manifest-with-the-yeoman-generator-for-office-add-ins)
- [<span data-ttu-id="abef8-108">Usar o log de tempo de execução para depurar seu suplemento</span><span class="sxs-lookup"><span data-stu-id="abef8-108">Use runtime logging to debug your add-in</span></span>](#use-runtime-logging-to-debug-your-add-in)


## <a name="validate-your-manifest-with-the-office-add-in-validator"></a><span data-ttu-id="abef8-109">Validar o manifesto com o Validador de Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="abef8-109">Validate your manifest with the Office Add-in Validator</span></span>

<span data-ttu-id="abef8-110">Para ajudar a garantir que o arquivo de manifesto que descreve o suplemento do Office está correto e completo, valide-o com base no [Validador de Suplemento do Office](https://github.com/OfficeDev/office-addin-validator).</span><span class="sxs-lookup"><span data-stu-id="abef8-110">To help ensure that the manifest file that describes your Office Add-in is correct and complete, validate it against the [Office Add-in Validator](https://github.com/OfficeDev/office-addin-validator).</span></span>

### <a name="to-use-the-office-add-in-validator-to-validate-your-manifest"></a><span data-ttu-id="abef8-111">Para usar o Validador de Suplemento do Office para validar o manifesto:</span><span class="sxs-lookup"><span data-stu-id="abef8-111">To use the Office Add-in Validator to validate your manifest</span></span>

1. <span data-ttu-id="abef8-112">Instale o [Node.js](https://nodejs.org/download/).</span><span class="sxs-lookup"><span data-stu-id="abef8-112">Install [Node.js](https://nodejs.org/download/).</span></span> 

2. <span data-ttu-id="abef8-113">Abra um prompt de comando/terminal como administrador e instale o Validador de Suplemento do Office e as respectivas dependências globalmente usando o seguinte comando:</span><span class="sxs-lookup"><span data-stu-id="abef8-113">Open a command prompt / terminal as an administrator, and install the Office Add-in Validator and its dependencies globally by using the following command:</span></span>

    ```bash
    npm install -g office-addin-validator
    ```
    
    > [!NOTE]
    > <span data-ttu-id="abef8-114">se já instalou o Office, atualize para a versão mais recente para que o validador seja instalado como uma dependência.</span><span class="sxs-lookup"><span data-stu-id="abef8-114">If you already have Yo Office installed, upgrade to the latest version, and the validator will be installed as a dependency.</span></span>

3. <span data-ttu-id="abef8-p101">Para validar o manifesto, execute o seguinte comando: substitua MANIFEST.XML pelo caminho para o arquivo XML de manifesto.</span><span class="sxs-lookup"><span data-stu-id="abef8-p101">Run the following command to validate your manifest. Replace MANIFEST.XML with the path to the manifest XML file.</span></span>

    ```bash
    validate-office-addin MANIFEST.XML
    ```

## <a name="validate-your-manifest-against-the-xml-schema"></a><span data-ttu-id="abef8-117">Validar seu manifesto em relação ao esquema XML</span><span class="sxs-lookup"><span data-stu-id="abef8-117">Validate your manifest against the XML schema</span></span>

<span data-ttu-id="abef8-118">Para ajudar a garantir que o arquivo de manifesto segue o esquema correto, incluindo todos os namespaces de elementos que você está usando.</span><span class="sxs-lookup"><span data-stu-id="abef8-118">To help ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using.</span></span> <span data-ttu-id="abef8-119">Se você copiou elementos de outros manifestos da amostra, verifique se também **incluiu os namespaces apropriados**.</span><span class="sxs-lookup"><span data-stu-id="abef8-119">If you copied elements from other sample manifests double check you also **include the appropriate namespaces**.</span></span> <span data-ttu-id="abef8-120">É possível validar um manifesto em relação aos arquivos de [Definição de esquema XML (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas).</span><span class="sxs-lookup"><span data-stu-id="abef8-120">You can validate a manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) files.</span></span> <span data-ttu-id="abef8-121">É possível usar uma ferramenta de validação de esquema XML para executar essa validação.</span><span class="sxs-lookup"><span data-stu-id="abef8-121">You can use an XML schema validation tool to perform this validation.</span></span> 



### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a><span data-ttu-id="abef8-122">Para usar uma ferramenta de validação de esquema XML da linha de comando para validar seu manifesto</span><span class="sxs-lookup"><span data-stu-id="abef8-122">To use a command-line XML schema validation tool to validate your manifest</span></span>

1.  <span data-ttu-id="abef8-123">Instale o [tar](https://www.gnu.org/software/tar/) e o [libxml](http://xmlsoft.org/FAQ.html), caso ainda não tenha o feito.</span><span class="sxs-lookup"><span data-stu-id="abef8-123">Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.</span></span>

2.  <span data-ttu-id="abef8-p103">Execute o comando a seguir. Substitua `XSD_FILE` pelo caminho para o arquivo XSD do manifesto e `XML_FILE` pelo caminho para o arquivo XML do manifesto.</span><span class="sxs-lookup"><span data-stu-id="abef8-p103">Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.</span></span>
    
    ```bash
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a><span data-ttu-id="abef8-126">Validar o manifesto com o gerador Yeoman para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="abef8-126">Validate your manifest with the Yeoman generator for Office Add-ins</span></span>

<span data-ttu-id="abef8-127">Caso tenha criado o Suplemento do Office usando o [Gerador Yeoman para Suplementos do Office](https://www.npmjs.com/package/generator-office), é possível garantir que o arquivo de manifesto segue o esquema correto executando o seguinte comando no diretório raiz do projeto:</span><span class="sxs-lookup"><span data-stu-id="abef8-127">If you've created your Office Add-in using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office), you can ensure that the manifest file follows the correct schema by running the following command within the root directory of your project:</span></span>

```bash
npm run validate
```

![Gif animado que mostra o validador Yo Office em execução na linha de comando e gerando os resultados que mostram que a validação foi aprovada](../images/yo-office-validator.gif)

> [!NOTE]
> <span data-ttu-id="abef8-129">Para ter acesso a essa funcionalidade, o projeto de suplemento deve ter sido criado usando o [Gerador Yeoman para Suplementos do Office](https://www.npmjs.com/package/generator-office) versão 1.1.17 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="abef8-129">To have access to this functionality, your add-in project must have been created by using [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) version 1.1.17 or later.</span></span>

## <a name="use-runtime-logging-to-debug-your-add-in"></a><span data-ttu-id="abef8-130">Usar o log de tempo de execução para depurar seu suplemento</span><span class="sxs-lookup"><span data-stu-id="abef8-130">Use runtime logging to debug your add-in</span></span> 

<span data-ttu-id="abef8-131">Você pode usar o log de tempo de execução para depurar o manifesto do seu suplemento, assim como diversos erros de instalação.</span><span class="sxs-lookup"><span data-stu-id="abef8-131">You can use runtime logging to debug your add-in's manifest as well as several installation errors.</span></span> <span data-ttu-id="abef8-132">Esse recurso pode ajudá-lo a identificar e corrigir problemas com seu manifesto que não são detectados pela validação de esquema XSD, como uma incompatibilidade entre as identificações dos recursos.</span><span class="sxs-lookup"><span data-stu-id="abef8-132">This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs.</span></span> <span data-ttu-id="abef8-133">O log de tempo de execução é particularmente útil para depurar suplementos que implementam comandos de suplemento e funções personalizadas do Excel.</span><span class="sxs-lookup"><span data-stu-id="abef8-133">Runtime logging is particularly  useful for debugging add-ins that implement add-in commands and Excel custom functions.</span></span>   

> [!NOTE]
> <span data-ttu-id="abef8-134">O recurso de log de tempo de execução está atualmente disponível para o Office 2016 para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="abef8-134">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

### <a name="to-turn-on-runtime-logging"></a><span data-ttu-id="abef8-135">Para ativar o log de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="abef8-135">To turn on runtime logging</span></span>

> [!IMPORTANT]
> <span data-ttu-id="abef8-p105">O log do tempo de execução afeta o desempenho. Ative-o somente quando precisar depurar problemas com seu manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="abef8-p105">Runtime Logging affects performance. Turn it on only when you need to debug issues with your add-in manifest.</span></span>

<span data-ttu-id="abef8-138">Para ativar o log de tempo de execução:</span><span class="sxs-lookup"><span data-stu-id="abef8-138">To turn on runtime logging:</span></span>

1. <span data-ttu-id="abef8-139">Verifique se você está executando o Office 2016 para área de trabalho na compilação **16.0.7019** ou posterior.</span><span class="sxs-lookup"><span data-stu-id="abef8-139">Make sure that you are running Office 2016 desktop build **16.0.7019** or later.</span></span> 

2. <span data-ttu-id="abef8-140">Adicione a chave do registro `RuntimeLogging` em `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`</span><span class="sxs-lookup"><span data-stu-id="abef8-140">Add the `RuntimeLogging` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="abef8-141">Se a chave (pasta) `Developer` ainda não existir em `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, conclua as seguintes etapas para criá-la:</span><span class="sxs-lookup"><span data-stu-id="abef8-141">If the `Developer` key (folder) does not already exist under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, complete the following steps to create it:</span></span> 
    > 1. <span data-ttu-id="abef8-142">Clique com o botão direito do mouse na chave (pasta) **WEF** e selecione **Novo** > **Chave**.</span><span class="sxs-lookup"><span data-stu-id="abef8-142">Right-click the **WEF** key (folder) and select **New** > **Key**.</span></span>
    > 2. <span data-ttu-id="abef8-143">Nomeie a nova chave como **Developer**.</span><span class="sxs-lookup"><span data-stu-id="abef8-143">Name the new key **Developer**.</span></span>

3. <span data-ttu-id="abef8-p106">Defina o valor padrão da chave para o caminho completo do arquivo onde você deseja que o log seja gravado. Para obter um exemplo, veja [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span><span class="sxs-lookup"><span data-stu-id="abef8-p106">Set the default value of the key to the full path of the file where you want the log to be written. For an example, see [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="abef8-146">A pasta na qual o arquivo de log será gravado deverá existir e você precisará ter permissões de gravação.</span><span class="sxs-lookup"><span data-stu-id="abef8-146">The directory in which the log file will be written must already exist, and you must have write permissions to it.</span></span> 
 
<span data-ttu-id="abef8-p107">A imagem a seguir mostra qual deve ser a aparência do registro. Para desativar o recurso, remova a chave do registro `RuntimeLogging`.</span><span class="sxs-lookup"><span data-stu-id="abef8-p107">The following image shows what the registry should look like. To turn the feature off, remove the `RuntimeLogging` key from the registry.</span></span> 

![Captura de tela do editor do registro com uma chave de registro RuntimeLogging](http://i.imgur.com/Sa9TyI6.png)


### <a name="to-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="abef8-150">Para solucionar problemas com o manifesto</span><span class="sxs-lookup"><span data-stu-id="abef8-150">To troubleshoot issues with your manifest</span></span>

<span data-ttu-id="abef8-151">Para usar o log do tempo de execução para solucionar problemas ao carregar um suplemento:</span><span class="sxs-lookup"><span data-stu-id="abef8-151">To use runtime logging to troubleshoot issues loading an add-in:</span></span>
 
1. <span data-ttu-id="abef8-152">[Realize o sideload do seu suplemento](sideload-office-add-ins-for-testing.md) para teste.</span><span class="sxs-lookup"><span data-stu-id="abef8-152">[Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="abef8-153">Recomendamos realizar o sideload apenas do suplemento que você está testando para minimizar a quantidade de mensagens no arquivo de log.</span><span class="sxs-lookup"><span data-stu-id="abef8-153">We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.</span></span>

2. <span data-ttu-id="abef8-154">Se nada acontecer e você não vir seu suplemento (e ele não estiver aparecendo na caixa de diálogo de suplementos), abra o arquivo de log.</span><span class="sxs-lookup"><span data-stu-id="abef8-154">If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.</span></span>

3. <span data-ttu-id="abef8-p108">Procure pela ID de seu suplemento no arquivo de log, definida no seu manifesto. No arquivo de log, essa ID está marcada como `SolutionId`.</span><span class="sxs-lookup"><span data-stu-id="abef8-p108">Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`.</span></span> 

<span data-ttu-id="abef8-p109">No exemplo a seguir, o arquivo de log identifica um controle que aponta para um arquivo de recurso que não existe. Neste exemplo, a correção seria reparar o erro de digitação no manifesto ou adicionar o recurso que está faltando.</span><span class="sxs-lookup"><span data-stu-id="abef8-p109">In the following example, the log file identifies a control that points to a resource file that doesn't exist. For this example, the fix would be to correct the typo in the manifest or to add the missing resource.</span></span>

![Captura de tela de um arquivo de log com uma entrada que especifica uma identificação de recurso que não foi encontrado](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a><span data-ttu-id="abef8-160">Problemas conhecidos com o log de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="abef8-160">Known issues with runtime logging</span></span>

<span data-ttu-id="abef8-p110">Talvez você veja mensagens no arquivo de log que são confusas ou que estão classificadas incorretamente. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="abef8-p110">You might see messages in the log file that are confusing or that are classified incorrectly. For example:</span></span>

- <span data-ttu-id="abef8-163">A mensagem `Medium Current host not in add-in's host list` seguida por `Unexpected Parsed manifest targeting different host` é incorretamente classificada como um erro.</span><span class="sxs-lookup"><span data-stu-id="abef8-163">The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.</span></span>

- <span data-ttu-id="abef8-164">Se você vir a mensagem `Unexpected Add-in is missing required manifest fields DisplayName` e ela não contiver uma SolutionId, o erro provavelmente não está relacionado ao suplemento que você está depurando.</span><span class="sxs-lookup"><span data-stu-id="abef8-164">If you see the message `Unexpected Add-in is missing required manifest fields DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging.</span></span> 

- <span data-ttu-id="abef8-p111">Todas as mensagens `Monitorable` indicam erros esperados do ponto de vista do sistema. Às vezes, indicam um problema com o seu manifesto, como um elemento que foi soletrado incorretamente e que foi ignorado, mas que não fez com que o manifesto falhasse.</span><span class="sxs-lookup"><span data-stu-id="abef8-p111">Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail.</span></span> 

## <a name="clear-the-office-cache"></a><span data-ttu-id="abef8-167">Limpar o cache do Office</span><span class="sxs-lookup"><span data-stu-id="abef8-167">Clear the Office cache</span></span>

<span data-ttu-id="abef8-168">Se parecer que as alterações que você fez no manifesto, como nomes de arquivo dos ícones de botão da faixa de opções ou o texto de comandos de suplemento, não entraram em vigor, tente limpar o cache do Office no computador.</span><span class="sxs-lookup"><span data-stu-id="abef8-168">If changes you've made in the manifest, such as file names of ribbon button icons, or text of add-in commands, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="abef8-169">No Windows:</span><span class="sxs-lookup"><span data-stu-id="abef8-169">For Windows:</span></span>
<span data-ttu-id="abef8-170">Exclua o conteúdo da pasta `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="abef8-170">Delete the content of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="abef8-171">No Mac:</span><span class="sxs-lookup"><span data-stu-id="abef8-171">For Mac:</span></span>
<span data-ttu-id="abef8-172">Exclua o conteúdo da pasta `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="abef8-172">Delete the content of the folder `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>

#### <a name="for-ios"></a><span data-ttu-id="abef8-173">No iOS:</span><span class="sxs-lookup"><span data-stu-id="abef8-173">For iOS:</span></span>
<span data-ttu-id="abef8-p112">Chame `window.location.reload(true)` usando o JavaScript no suplemento para forçar um recarregamento. Outra alternativa é reinstalar o Office.</span><span class="sxs-lookup"><span data-stu-id="abef8-p112">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="abef8-176">Confira também</span><span class="sxs-lookup"><span data-stu-id="abef8-176">See also</span></span>

- [<span data-ttu-id="abef8-177">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="abef8-177">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="abef8-178">Realizar sideload de suplementos do Office para teste</span><span class="sxs-lookup"><span data-stu-id="abef8-178">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="abef8-179">Depurar suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="abef8-179">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
