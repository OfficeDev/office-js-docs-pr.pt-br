---
title: Depurar seu suplemento com o log de tempo de execução
description: Saiba como usar o log do tempo de execução para depurar seu suplemento.
ms.date: 12/31/2019
localization_priority: Normal
ms.openlocfilehash: c8d58a36c4203708c87cfc9b6c6b1790f05876a7
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950919"
---
# <a name="debug-your-add-in-with-runtime-logging"></a><span data-ttu-id="17181-103">Depurar seu suplemento com o log de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="17181-103">Debug your add-in with runtime logging</span></span>

<span data-ttu-id="17181-104">Você pode usar o log de tempo de execução para depurar o manifesto do seu suplemento, assim como diversos erros de instalação.</span><span class="sxs-lookup"><span data-stu-id="17181-104">You can use runtime logging to debug your add-in's manifest as well as several installation errors.</span></span> <span data-ttu-id="17181-105">Esse recurso pode ajudá-lo a identificar e corrigir problemas com seu manifesto que não são detectados pela validação de esquema XSD, como uma incompatibilidade entre as identificações dos recursos.</span><span class="sxs-lookup"><span data-stu-id="17181-105">This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs.</span></span> <span data-ttu-id="17181-106">O log de tempo de execução é particularmente útil para depurar suplementos que implementam comandos de suplemento e funções personalizadas do Excel.</span><span class="sxs-lookup"><span data-stu-id="17181-106">Runtime logging is particularly  useful for debugging add-ins that implement add-in commands and Excel custom functions.</span></span>   

> [!NOTE]
> <span data-ttu-id="17181-107">O recurso de log de tempo de execução está atualmente disponível para o Office 2016 para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="17181-107">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="17181-108">o log do tempo de execução afeta o desempenho.</span><span class="sxs-lookup"><span data-stu-id="17181-108">Runtime Logging affects performance.</span></span> <span data-ttu-id="17181-109">Ative-o somente quando precisar depurar problemas com o manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="17181-109">Turn it on only when you need to debug issues with your add-in manifest.</span></span>

## <a name="use-runtime-logging-from-the-command-line"></a><span data-ttu-id="17181-110">Use o log de tempo de execução na linha de comandos</span><span class="sxs-lookup"><span data-stu-id="17181-110">Use runtime logging from the command line</span></span>

<span data-ttu-id="17181-111">Habilitar o log de tempo de execução na linha de comando é a maneira mais rápida de usar essa ferramenta de log.</span><span class="sxs-lookup"><span data-stu-id="17181-111">Enabling runtime logging from the command line is the fastest way to use this logging tool.</span></span> <span data-ttu-id="17181-112">Esse usa npx, que é fornecido por padrão como parte do npm@5.2.0+.</span><span class="sxs-lookup"><span data-stu-id="17181-112">These use npx, which is provided by default as part of npm@5.2.0+.</span></span> <span data-ttu-id="17181-113">Se você possui uma versão anterior do [npm](https://www.npmjs.com/), tente as instruções do [Log de tempo de execução no Windows](#runtime-logging-on-windows) ou do [Log de tempo de execução no Mac](#runtime-logging-on-mac), ou [instale o npx](https://www.npmjs.com/package/npx).</span><span class="sxs-lookup"><span data-stu-id="17181-113">If you have an earlier version of [npm](https://www.npmjs.com/), try [Runtime logging on Windows](#runtime-logging-on-windows) or [Runtime logging on Mac](#runtime-logging-on-mac) instructions, or [install npx](https://www.npmjs.com/package/npx).</span></span>

- <span data-ttu-id="17181-114">Para habilitar o log de tempo de execução:</span><span class="sxs-lookup"><span data-stu-id="17181-114">To enable runtime logging:</span></span>
    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable
    ```
- <span data-ttu-id="17181-115">Para habilitar o log de tempo de execução apenas para um arquivo específico, use o mesmo comando com um nome de arquivo:</span><span class="sxs-lookup"><span data-stu-id="17181-115">To enable runtime logging only for a specific file, use the same command with a filename:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable [filename.txt]
    ```

- <span data-ttu-id="17181-116">Para desabilitar o log de tempo de execução:</span><span class="sxs-lookup"><span data-stu-id="17181-116">To disable runtime logging:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --disable
    ```

- <span data-ttu-id="17181-117">Para exibir se o log de tempo de execução está ativado:</span><span class="sxs-lookup"><span data-stu-id="17181-117">To display whether runtime logging is enabled:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log
    ```

- <span data-ttu-id="17181-118">Para exibir ajuda na linha de comandos para o log de tempo de execução:</span><span class="sxs-lookup"><span data-stu-id="17181-118">To display help within the command line for runtime logging:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --help
    ```

## <a name="runtime-logging-on-windows"></a><span data-ttu-id="17181-119">Log de tempo de execução no Windows</span><span class="sxs-lookup"><span data-stu-id="17181-119">Runtime logging on Windows</span></span>

1. <span data-ttu-id="17181-120">Verifique se você está executando o Office 2016 para área de trabalho na compilação **16.0.7019** ou posterior.</span><span class="sxs-lookup"><span data-stu-id="17181-120">Make sure that you are running Office 2016 desktop build **16.0.7019** or later.</span></span> 

2. <span data-ttu-id="17181-121">Adicione a chave do registro `RuntimeLogging` em `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`</span><span class="sxs-lookup"><span data-stu-id="17181-121">Add the `RuntimeLogging` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="17181-122">Se a chave (pasta) `Developer` ainda não existir em `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, conclua as seguintes etapas para criá-la:</span><span class="sxs-lookup"><span data-stu-id="17181-122">If the `Developer` key (folder) does not already exist under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, complete the following steps to create it:</span></span> 
    > 1. <span data-ttu-id="17181-123">Clique com o botão direito do mouse na chave (pasta) **WEF** e selecione **Novo** > **Chave**.</span><span class="sxs-lookup"><span data-stu-id="17181-123">Right-click the **WEF** key (folder) and select **New** > **Key**.</span></span>
    > 2. <span data-ttu-id="17181-124">Nomeie a nova chave como **Developer**.</span><span class="sxs-lookup"><span data-stu-id="17181-124">Name the new key **Developer**.</span></span>

3. <span data-ttu-id="17181-125">Defina o valor padrão da chave **RuntimeLogging** para o caminho completo do arquivo em que você deseja que o log seja gravado.</span><span class="sxs-lookup"><span data-stu-id="17181-125">Set the default value of the **RuntimeLogging** key to the full path of the file where you want the log to be written.</span></span> <span data-ttu-id="17181-126">Para obter um exemplo, veja [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span><span class="sxs-lookup"><span data-stu-id="17181-126">For an example, see [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="17181-127">A pasta na qual o arquivo de log será gravado deverá existir e você precisará ter permissões de gravação.</span><span class="sxs-lookup"><span data-stu-id="17181-127">The directory in which the log file will be written must already exist, and you must have write permissions to it.</span></span> 
 
<span data-ttu-id="17181-p105">A imagem a seguir mostra qual deve ser a aparência do registro. Para desativar o recurso, remova a chave do registro `RuntimeLogging`.</span><span class="sxs-lookup"><span data-stu-id="17181-p105">The following image shows what the registry should look like. To turn the feature off, remove the `RuntimeLogging` key from the registry.</span></span> 

![Captura de tela do editor do registro com uma chave de registro RuntimeLogging](http://i.imgur.com/Sa9TyI6.png)

## <a name="runtime-logging-on-mac"></a><span data-ttu-id="17181-131">Log de tempo de execução no Mac</span><span class="sxs-lookup"><span data-stu-id="17181-131">Runtime logging on Mac</span></span>

1. <span data-ttu-id="17181-132">Verifique se você está executando o build de área de trabalho do Office 2016 **16.27** (19071500) ou posterior.</span><span class="sxs-lookup"><span data-stu-id="17181-132">Make sure that you are running Office 2016 desktop build **16.27** (19071500) or later.</span></span>

2. <span data-ttu-id="17181-133">Abra o **Terminal** e defina uma preferência de log de tempo de execução usando o comando `defaults`:</span><span class="sxs-lookup"><span data-stu-id="17181-133">Open **Terminal** and set a runtime logging preference by using the `defaults` command:</span></span>
    
    ```command&nbsp;line
    defaults write <bundle id> CEFRuntimeLoggingFile -string <file_name>
    ```

    <span data-ttu-id="17181-134">`<bundle id>` identifica quais hosts devem ser habilitados no log de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="17181-134">`<bundle id>` identifies which the host for which to enable runtime logging.</span></span> <span data-ttu-id="17181-135">`<file_name>` é o nome do arquivo de texto no qual o log será gravado.</span><span class="sxs-lookup"><span data-stu-id="17181-135">`<file_name>` is the name of the text file to which the log will be written.</span></span>

    <span data-ttu-id="17181-136">Defina `<bundle id>` para um dos seguintes valores para habilitar o log de tempo de execução do host correspondente:</span><span class="sxs-lookup"><span data-stu-id="17181-136">Set `<bundle id>` to one of the following values to enable runtime logging for the corresponding host:</span></span>

    - `com.microsoft.Word`
    - `com.microsoft.Excel`
    - `com.microsoft.Powerpoint`
    - `com.microsoft.Outlook`

<span data-ttu-id="17181-137">O exemplo a seguir habilita o log de tempo de execução do Word e, em seguida, abre o arquivo de log:</span><span class="sxs-lookup"><span data-stu-id="17181-137">The following example enables runtime logging for Word and then opens the log file:</span></span>

```command&nbsp;line
defaults write com.microsoft.Word CEFRuntimeLoggingFile -string "runtime_logs.txt"
open ~/library/Containers/com.microsoft.Word/Data/runtime_logs.txt
```

> [!NOTE] 
> <span data-ttu-id="17181-138">Será preciso reiniciar o Office depois de executar o comando `defaults` para habilitar o log de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="17181-138">You'll need to restart Office after running the `defaults` command to enable runtime logging.</span></span>

<span data-ttu-id="17181-139">Para desativar o log de tempo de execução, use o comando `defaults delete`:</span><span class="sxs-lookup"><span data-stu-id="17181-139">To turn off runtime logging, use the `defaults delete` command:</span></span>

```command&nbsp;line
defaults delete <bundle id> CEFRuntimeLoggingFile
```

<span data-ttu-id="17181-140">O exemplo a seguir desabilitará o log de tempo de execução do Word:</span><span class="sxs-lookup"><span data-stu-id="17181-140">The following example will turn off runtime logging for Word:</span></span>

```command&nbsp;line
defaults delete com.microsoft.Word CEFRuntimeLoggingFile
```

## <a name="use-runtime-logging-to-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="17181-141">Use o log do tempo de execução para solucionar problemas em seu manifesto</span><span class="sxs-lookup"><span data-stu-id="17181-141">Use runtime logging to troubleshoot issues with your manifest</span></span>

<span data-ttu-id="17181-142">Para usar o log do tempo de execução para solucionar problemas ao carregar um suplemento:</span><span class="sxs-lookup"><span data-stu-id="17181-142">To use runtime logging to troubleshoot issues loading an add-in:</span></span>
 
1. <span data-ttu-id="17181-143">[Realize o sideload do seu suplemento](sideload-office-add-ins-for-testing.md) para teste.</span><span class="sxs-lookup"><span data-stu-id="17181-143">[Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="17181-144">Recomendamos realizar o sideload apenas do suplemento que você está testando para minimizar a quantidade de mensagens no arquivo de log.</span><span class="sxs-lookup"><span data-stu-id="17181-144">We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.</span></span>

2. <span data-ttu-id="17181-145">Se nada acontecer e você não vir seu suplemento (e ele não estiver aparecendo na caixa de diálogo de suplementos), abra o arquivo de log.</span><span class="sxs-lookup"><span data-stu-id="17181-145">If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.</span></span>

3. <span data-ttu-id="17181-p107">Procure pela ID de seu suplemento no arquivo de log, definida no seu manifesto. No arquivo de log, essa ID está marcada como `SolutionId`.</span><span class="sxs-lookup"><span data-stu-id="17181-p107">Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`.</span></span> 

<span data-ttu-id="17181-p108">No exemplo a seguir, o arquivo de log identifica um controle que aponta para um arquivo de recurso que não existe. Neste exemplo, a correção seria reparar o erro de digitação no manifesto ou adicionar o recurso que está faltando.</span><span class="sxs-lookup"><span data-stu-id="17181-p108">In the following example, the log file identifies a control that points to a resource file that doesn't exist. For this example, the fix would be to correct the typo in the manifest or to add the missing resource.</span></span>

![Captura de tela de um arquivo de log com uma entrada que especifica uma identificação de recurso que não foi encontrado](http://i.imgur.com/f8bouLA.png) 

## <a name="known-issues-with-runtime-logging"></a><span data-ttu-id="17181-151">Problemas conhecidos com o log de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="17181-151">Known issues with runtime logging</span></span>

<span data-ttu-id="17181-p109">Talvez você veja mensagens no arquivo de log que são confusas ou que estão classificadas incorretamente. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="17181-p109">You might see messages in the log file that are confusing or that are classified incorrectly. For example:</span></span>

- <span data-ttu-id="17181-154">A mensagem `Medium Current host not in add-in's host list` seguida por `Unexpected Parsed manifest targeting different host` é incorretamente classificada como um erro.</span><span class="sxs-lookup"><span data-stu-id="17181-154">The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.</span></span>

- <span data-ttu-id="17181-155">Se você vir a mensagem `Unexpected Add-in is missing required manifest fields DisplayName` e ela não contiver uma SolutionId, o erro provavelmente não está relacionado ao suplemento que você está depurando.</span><span class="sxs-lookup"><span data-stu-id="17181-155">If you see the message `Unexpected Add-in is missing required manifest fields DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging.</span></span> 

- <span data-ttu-id="17181-p110">Todas as mensagens `Monitorable` indicam erros esperados do ponto de vista do sistema. Às vezes, indicam um problema com o seu manifesto, como um elemento que foi soletrado incorretamente e que foi ignorado, mas que não fez com que o manifesto falhasse.</span><span class="sxs-lookup"><span data-stu-id="17181-p110">Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail.</span></span> 

## <a name="see-also"></a><span data-ttu-id="17181-158">Confira também</span><span class="sxs-lookup"><span data-stu-id="17181-158">See also</span></span>

- [<span data-ttu-id="17181-159">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="17181-159">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="17181-160">Validar o manifesto de suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="17181-160">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="17181-161">Limpar o cache do Office</span><span class="sxs-lookup"><span data-stu-id="17181-161">Clear the Office cache</span></span>](clear-cache.md)
- [<span data-ttu-id="17181-162">Realizar sideload de suplementos do Office para teste</span><span class="sxs-lookup"><span data-stu-id="17181-162">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="17181-163">Depurar suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="17181-163">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)