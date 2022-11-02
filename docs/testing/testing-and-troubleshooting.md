---
title: Solucionar erros de usuários com suplementos do Office
description: Saiba como solucionar problemas de erros do usuário em Suplementos do Office.
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 18bb3c180cd3af1eb8d045d7c69b9772532b04d4
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810369"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a>Solucionar erros de usuários com suplementos do Office

At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.

Também é possível usar o [Fiddler](https://www.telerik.com/fiddler) para identificar e depurar problemas com os suplementos.

## <a name="common-errors-and-troubleshooting-steps"></a>Erros comuns e etapas de solução de problemas

A tabela a seguir lista as mensagens de erro comuns que os usuários podem receber e as etapas que os usuários podem seguir para resolver os erros.

|**Mensagem de erro**|**Resolução**|
|:-----|:-----|
|Erro do aplicativo: catálogo não pôde ser alcançado|Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.|
|APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.|Verifique se as atualizações mais recentes do Office foram instaladas ou baixe a [atualização do Office 2013](https://support.microsoft.com/kb/2986156/).|
|Erro: objeto não dá suporte à propriedade ou ao método 'defineProperty'|Confirme se o Internet Explorer não está sendo executado no modo de compatibilidade. Acesse **Configurações de exibição de compatibilidade** **de ferramentas** > .|
|Não foi possível carregar o aplicativo porque não há suporte para sua versão do navegador. Clique aqui para obter uma lista de versões do navegador compatíveis.|Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).|

## <a name="when-installing-an-add-in-you-see-error-loading-add-in-in-the-status-bar"></a>Ao instalar um suplemento, você verá “erro ao carregar suplemento” na barra de status

1. Feche o Office.
1. Verifique se o manifesto é válido. Consulte [Validar um manifesto do Suplemento do Office](troubleshoot-manifest.md).
1. Reiniciar o suplemento.
1. Instale o suplemento novamente.

Você também pode enviar comentários: se estiver usando o Excel no Windows ou Mac, envie comentários à equipe de extensibilidade do Office diretamente do Excel. Para fazer isso, selecione **Arquivo** > **Comentário** > **Enviar um Rosto Triste**. Enviando um rosto triste, você fornece os logs necessários para entendermos o problema.

## <a name="outlook-add-in-doesnt-work-correctly"></a>O suplemento do Outlook não funciona corretamente

Se um suplemento do Outlook executado no Windows e [usando o Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) não está funcionando corretamente, tente ativar a depuração de scripts no Internet Explorer.

- Acesse **Ferramentas****Opções** > **de Internet Avançadas** > .
- Em **Navegação**, desmarque **Desmarque a depuração de script (Internet Explorer)** e **Desabilite a depuração de script (Outros)**.

Recomendamos que você desmarque essas configurações somente para solucionar o problema. Se você deixar desmarcado, receberá prompts durante a navegação. Depois que o problema for resolvido, verifique **Desabilitar a depuração de script (Internet Explorer)** e **Desabilitar a depuração de script (Outros)** novamente.

## <a name="add-in-doesnt-activate-in-office-2013"></a>O suplemento não é ativado no Office 2013

Se o suplemento não for ativado quando o usuário executar as etapas a seguir.

1. Entrar com a conta da Microsoft no Office 2013.

1. Habilitar a verificação de duas etapas para a conta da Microsoft.

1. Verificar a identidade ao ser solicitado quando tentar inserir um suplemento.

Verifique se as atualizações mais recentes do Office foram instaladas ou baixe a [atualização do Office 2013](https://support.microsoft.com/kb/2986156/).

## <a name="add-in-dialog-box-cannot-be-displayed"></a>Não é possível exibir a caixa de diálogo do suplemento

Quando o usuário usa um suplemento do Office, ele é solicitado a permitir a exibição de uma caixa de diálogo. O usuário escolhe **Permitir** e a seguinte mensagem de erro ocorre.

"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."

![Captura de tela da mensagem de erro da caixa de diálogo.](../images/dialog-prevented.png)

|Navegadores afetados|Plataformas afetadas|
|:--------------------|:---------------------|
|Microsoft Edge|Office na Web|

Para resolver o problema, os usuários finais ou administradores podem adicionar o domínio do suplemento à lista de sites confiáveis no navegador Microsoft Edge.

> [!IMPORTANT]
> Caso não confie no suplemento, não adicione a respectiva URL à lista de sites confiáveis.

Para adicionar uma URL à lista de sites confiáveis:

1. No **Painel de Controle**, abra **Opções da Internet** > **Security**.
1. Escolha a zona de **Sites confiáveis** e escolha **Sites**.
1. Insira a URL exibida na mensagem de erro e escolha **Adicionar**.
1. Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.

This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInIFrame:true}, callback);
```

## <a name="see-also"></a>Veja também

- [Solucionar erros de desenvolvimento com Suplementos do Office](troubleshoot-development-errors.md)
