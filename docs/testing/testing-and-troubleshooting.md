---
title: Solucionar erros de usuários com suplementos do Office
description: ''
ms.date: 11/26/2019
localization_priority: Priority
ms.openlocfilehash: fed23cfde1380e7c5728c78c995d3b89d44451f2
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629732"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a>Solucionar erros de usuários com suplementos do Office

Às vezes, seus usuários podem encontrar problemas com suplementos do Office desenvolvidos por você. Por exemplo, um suplemento falha ao carregar ou está inacessível. Use as informações neste artigo para ajudar a resolver problemas comuns que os usuários têm com o seu suplemento do Office. 

Também é possível usar o [Fiddler](https://www.telerik.com/fiddler) para identificar e depurar problemas com os suplementos.

## <a name="common-errors-and-troubleshooting-steps"></a>Erros comuns e etapas de solução de problemas

A tabela a seguir lista as mensagens de erro comuns que os usuários podem receber e as etapas que os usuários podem seguir para resolver os erros.



|**Mensagem de erro**|**Resolução**|
|:-----|:-----|
|Erro do aplicativo: catálogo não pôde ser alcançado|Verifique as configurações do firewall. "Catálogo" refere-se ao AppSource. Essa mensagem indica que o usuário não consegue acessar o AppSource.|
|ERRO DO APLICATIVO: este aplicativo não pôde ser iniciado. Feche essa caixa de diálogo para ignorar o problema ou clique em "Reiniciar"para tentar novamente.|Verifique se as atualizações mais recentes do Office foram instaladas ou baixe a [atualização do Office 2013](https://support.microsoft.com/kb/2986156/).|
|Erro: objeto não dá suporte à propriedade ou ao método 'defineProperty'|Confirme se o Internet Explorer não está sendo executado no modo de compatibilidade. Vá para Ferramentas >  **Configurações do Modo de Exibição de Compatibilidade**.|
|Não foi possível carregar o aplicativo porque não há suporte para sua versão do navegador. Clique aqui para obter uma lista de versões do navegador compatíveis.|Verifique se o navegador dá suporte a armazenamento local HTML5 ou redefina as configurações do Internet Explorer. Para saber mais sobre os navegadores compatíveis, confira [Requisitos para a execução de Suplementos do Office](../concepts/requirements-for-running-office-add-ins.md).|

## <a name="when-installing-an-add-in-you-see-error-loading-add-in-in-the-status-bar"></a>Ao instalar um suplemento, você verá “erro ao carregar suplemento” na barra de status

1. Feche o Office.
2. Verifique se o manifesto é valido
3. Reinicie o suplemento
4. Instale o suplemento novamente.

Você também pode enviar comentários: se estiver usando o Excel no Windows ou Mac, envie comentários à equipe de extensibilidade do Office diretamente do Excel. Para fazer isso, selecione **Arquivo** | **Comentário** | **Enviar um Rosto Triste**. Enviando um rosto triste, você fornece os logs necessários para entendermos o problema.

## <a name="outlook-add-in-doesnt-work-correctly"></a>O suplemento do Outlook não funciona corretamente

Se um suplemento do Outlook executado no Windows e [usando o Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) não está funcionando corretamente, tente ativar a depuração de scripts no Internet Explorer. 


- Vá para Ferramentas > **Opções da Internet** > **Avançado**.
    
- Em **Navegação**, desmarque **Desabilitar depuração de scripts (Internet Explorer)** e **Desabilitar depuração de scripts (Outros)**.
    
Recomendamos que você desmarque essas configurações somente para solucionar o problema. Se você deixar desmarcado, receberá prompts durante a navegação. Depois que o problema for resolvido, marque **Desabilitar depuração de scripts (Internet Explorer)** e **Desabilitar depuração de scripts (Outros)** novamente.


## <a name="add-in-doesnt-activate-in-office-2013"></a>O suplemento não é ativado no Office 2013

Se o suplemento não for ativado quando o usuário executar as seguintes etapas:


1. Entrar com a conta da Microsoft no Office 2013.
    
2. Habilitar a verificação de duas etapas para a conta da Microsoft.
    
3. Verificar a identidade ao ser solicitado quando tentar inserir um suplemento.
    
Verifique se as atualizações mais recentes do Office foram instaladas ou baixe a [atualização do Office 2013](https://support.microsoft.com/kb/2986156/).


## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>Não é possível carregar o suplemento no painel de tarefas ou outros problemas relacionados ao manifesto do suplemento

Confira [Validar e solucionar problemas com seu manifesto](troubleshoot-manifest.md) para depurar problemas do manifesto de suplemento.


## <a name="add-in-dialog-box-cannot-be-displayed"></a>Não é possível exibir a caixa de diálogo do suplemento

Quando o usuário usa um suplemento do Office, ele é solicitado a permitir a exibição de uma caixa de diálogo. O usuário escolhe **Permitir** e, em seguida, recebe a seguinte mensagem de erro:

"As configurações de segurança do navegador nos impedem de criar uma caixa de diálogo. Tente outro navegador ou configure o navegador para que a 'URL' e o domínio mostrado na barra de endereço estejam na mesma zona de segurança".

![Captura de tela da mensagem de erro na caixa de diálogo](http://i.imgur.com/3mqmlgE.png)

|**Navegadores afetados**|**Plataformas afetadas**|
|:--------------------|:---------------------|
|Internet Explorer, Microsoft Edge|Office na Web|

Para resolver o problema, os administradores ou usuários finais podem adicionar o domínio do suplemento à lista de sites confiáveis no Internet Explorer. Use o mesmo procedimento se estiver trabalhando com o navegador Internet Explorer ou Microsoft Edge.

> [!IMPORTANT]
> Caso não confie no suplemento, não adicione a respectiva URL à lista de sites confiáveis.

Para adicionar uma URL à lista de sites confiáveis:

1. No **Painel de Controle**, abra **Opções da Internet** > **Security**.
2. Escolha a zona de **Sites confiáveis** e escolha **Sites**.
3. Insira a URL exibida na mensagem de erro e escolha **Adicionar**.
4. Tente usar o suplemento novamente. Se o problema persistir, verifique as configurações de outras zonas de segurança e confira se o domínio do suplemento está na mesma zona que a URL exibida na barra de endereços do aplicativo do Office.

Esse problema ocorre quando a API da caixa de diálogo é usada no modo pop-up. Para evitar esse problema, use o sinalizador [displayInFrame](/javascript/api/office/office.ui). Isso requer que a página tenha suporte para exibição dentro de um iframe. O exemplo a seguir mostra como usar o sinalizador.

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>Alterações nos comandos de suplemento, incluindo botões da faixa de opções e itens de menu, não entram em vigor

Se alterações feitas no manifesto, como nomes de arquivo de ícones de botão da faixa de opções ou texto de comandos de suplemento, não parecerem entrar em vigor, experimente limpar o cache do Office no computador. 

#### <a name="for-windows"></a>Para Windows:
Exclua os conteúdos da pasta `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.

#### <a name="for-mac"></a>Para Mac:

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a>No iOS:
Chame `window.location.reload(true)` usando o JavaScript no suplemento para forçar um recarregamento. Outra alternativa é reinstalar o Office.

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a>Alterações em arquivos estáticos, como JavaScript, HTML e CSS, não entram em vigor

O navegador pode estar armazenando esses arquivos em cache. Para evitar isso, desative o cache do lado do cliente ao desenvolver. Os detalhes dependerão do tipo de servidor que você estiver usando. Na maioria dos casos, envolve adicionar determinados cabeçalhos às respostas HTTP. Sugerimos o seguinte conjunto:

- Controle de cache: "privado, sem cache, sem armazenamento"
- Pragma: "sem cache"
- Expira: "-1"

Para um exemplo de como fazer isso em um servidor Node.JS Express, confira [este arquivo app.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js). Para um exemplo em um projeto ASP.NET, confira [este arquivo cshtml](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/src/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).

Se o seu suplemento estiver hospedado no Servidor de Informações da Internet (IIS), você também poderá adicionar o seguinte ao web.config.

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

Se essas etapas não parecerem funcionar a princípio, talvez seja necessário limpar o cache do navegador. Faça isso através da interface do usuário do navegador. Às vezes, o cache do Microsoft Edge não é limpo com êxito quando você tenta limpá-lo na interface do usuário do Edge. Se isso acontecer, execute o seguinte comando em um prompt de comando do Windows.

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="see-also"></a>Confira também

- [Depurar suplementos no Office na Web](debug-add-ins-in-office-online.md) 
- [Realizar sideload de um suplemento do Office no iPad e no Mac](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Depurar suplementos do Office no iPad e no Mac](debug-office-add-ins-on-ipad-and-mac.md)  
- [Validar e solucionar problemas com seu manifesto](troubleshoot-manifest.md)
