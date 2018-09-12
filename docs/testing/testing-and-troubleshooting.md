---
title: Solucionar erros de usuários com suplementos do Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: c56485cff0248484b53974c2685827045bbb68eb
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944059"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a>Solucionar erros de usuários com suplementos do Office

Às vezes, seus usuários podem encontrar problemas com suplementos do Office desenvolvidos por você. Por exemplo, um suplemento falha ao carregar ou está inacessível. Use as informações neste artigo para ajudar a resolver problemas comuns que os usuários têm com o seu suplemento do Office. 

Também é possível usar o [Fiddler](http://www.telerik.com/fiddler) para identificar e depurar problemas com os suplementos.

Depois de resolver o problema do usuário, é possível [responder diretamente às avaliações dos clientes no AppSource](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).

## <a name="common-errors-and-troubleshooting-steps"></a>Erros comuns e etapas de solução de problemas

A tabela a seguir lista as mensagens de erro comuns que os usuários podem receber e as etapas que os usuários podem seguir para resolver os erros.



|**Mensagem de erro**|**Resolução**|
|:-----|:-----|
|Erro do aplicativo: catálogo não pôde ser alcançado|Verifique as configurações do firewall. "Catálogo" refere-se ao AppSource. Essa mensagem indica que o usuário não consegue acessar o AppSource.|
|ERRO DO APLICATIVO: este aplicativo não pôde ser iniciado. Feche essa caixa de diálogo para ignorar o problema ou clique em "Reiniciar"para tentar novamente.|Verifique se as atualizações mais recentes do Office foram instaladas ou baixe a [atualização do Office 2013](https://support.microsoft.com/kb/2986156/).|
|Erro: objeto não dá suporte à propriedade ou ao método 'defineProperty'|Confirme se o Internet Explorer não está sendo executado no modo de compatibilidade. Vá para Ferramentas >  **Configurações do Modo de Exibição de Compatibilidade**.|
|Não foi possível carregar o aplicativo porque não há suporte para sua versão do navegador. Clique aqui para obter uma lista de versões do navegador compatíveis.|Verifique se o navegador dá suporte a armazenamento local HTML5 ou redefina as configurações do Internet Explorer. Para saber mais sobre os navegadores compatíveis, confira [Requisitos para a execução de Suplementos do Office](../concepts/requirements-for-running-office-add-ins.md).|


## <a name="outlook-add-in-doesnt-work-correctly"></a>O suplemento do Outlook não funciona corretamente

Se um suplemento do Outlook executado no Windows não está funcionando corretamente, tente ativar a depuração de scripts no Internet Explorer. 


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
|Internet Explorer, Microsoft Edge|Office Online|

Para resolver o problema, os administradores ou usuários finais podem adicionar o domínio do suplemento à lista de sites confiáveis no Internet Explorer. Use o mesmo procedimento se estiver trabalhando com o navegador Internet Explorer ou Microsoft Edge.

> [!IMPORTANT]
> Caso não confie no suplemento, não adicione a respectiva URL à lista de sites confiáveis.

Para adicionar uma URL à lista de sites confiáveis:

1. No Internet Explorer, escolha o botão Ferramentas e vá para **Opções da Internet** > **Segurança**.
2. Escolha a zona de **Sites confiáveis** e escolha **Sites**.
3. Insira a URL exibida na mensagem de erro e escolha **Adicionar**.
4. Tente usar o suplemento novamente. Se o problema persistir, verifique as configurações de outras zonas de segurança e confira se o domínio do suplemento está na mesma zona que a URL exibida na barra de endereços do aplicativo do Office.

Esse problema ocorre quando a API da caixa de diálogo é usada no modo pop-up. Para evitar esse problema, use o sinalizador [displayInFrame](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js). Isso requer que a página tenha suporte para exibição dentro de um iframe. O exemplo a seguir mostra como usar o sinalizador.

```js

Office.context.ui.displayDialogAsync(startAddress, {displayInFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>Alterações nos comandos de suplemento, incluindo botões da faixa de opções e itens de menu, não entram em vigor
Às vezes, as alterações nos comandos de suplemento, como o ícone de um botão da faixa de opções ou o texto de um item de menu, não parecem entrar em vigor. Limpe o cache do Office das versões antigas.

#### <a name="for-windows"></a>No Windows:
Exclua o conteúdo da pasta `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.

#### <a name="for-mac"></a>No Mac:
Exclua o conteúdo da pasta `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.

#### <a name="for-ios"></a>No iOS:
Chame `window.location.reload(true)` usando o JavaScript no suplemento para forçar um recarregamento. Outra alternativa é reinstalar o Office.

## <a name="see-also"></a>Veja também

- [Depurar suplementos no Office Online](debug-add-ins-in-office-online.md) 
- [Realizar sideload de um suplemento do Office no iPad e no Mac](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Depurar suplementos do Office no iPad e no Mac](debug-office-add-ins-on-ipad-and-mac.md)  
- [Validar e solucionar problemas com seu manifesto](troubleshoot-manifest.md)
    
