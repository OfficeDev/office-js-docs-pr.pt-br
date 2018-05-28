---
title: Solucionar erros de usu?rios com suplementos do Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 375b3819d423362c7d5e124700a0bea2dcf6e9e0
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a>Solucionar erros de usu?rios com suplementos do Office

?s vezes, seus usu?rios podem encontrar problemas com suplementos do Office desenvolvidos por voc?. Por exemplo, um suplemento falha ao carregar ou est? inacess?vel. Use as informa??es neste artigo para ajudar a resolver problemas comuns que os usu?rios t?m com o seu suplemento do Office. 

Tamb?m ? poss?vel usar o [Fiddler](http://www.telerik.com/fiddler) para identificar e depurar problemas com os suplementos.

Depois de resolver o problema do usu?rio, ? poss?vel [responder diretamente ?s avalia??es dos clientes no AppSource](https://docs.microsoft.com/en-us/office/dev/store/create-effective-office-store-listings).

## <a name="common-errors-and-troubleshooting-steps"></a>Erros comuns e etapas de solu??o de problemas

A tabela a seguir lista as mensagens de erro comuns que os usu?rios podem receber e as etapas que os usu?rios podem seguir para resolver os erros.



|**Mensagem de erro**|**Resolu??o**|
|:-----|:-----|
|Erro do aplicativo: cat?logo n?o p?de ser alcan?ado|Verifique as configura??es do firewall. "Cat?logo" refere-se ao AppSource. Essa mensagem indica que o usu?rio n?o consegue acessar o AppSource.|
|ERRO DO APLICATIVO: este aplicativo n?o p?de ser iniciado. Feche essa caixa de di?logo para ignorar o problema ou clique em "Reiniciar"para tentar novamente.|Verifique se as atualiza??es mais recentes do Office foram instaladas ou baixe a [atualiza??o do Office 2013](https://support.microsoft.com/en-us/kb/2986156/).|
|Erro: objeto n?o d? suporte ? propriedade ou ao m?todo 'defineProperty'|Confirme se o Internet Explorer n?o est? sendo executado no modo de compatibilidade. V? para Ferramentas >  **Configura??es do Modo de Exibi??o de Compatibilidade**.|
|N?o foi poss?vel carregar o aplicativo porque n?o h? suporte para sua vers?o do navegador. Clique aqui para obter uma lista de vers?es do navegador compat?veis.|Verifique se o navegador d? suporte a armazenamento local HTML5 ou redefina as configura??es do Internet Explorer. Para saber mais sobre os navegadores compat?veis, confira [Requisitos para a execu??o de Suplementos do Office](../concepts/requirements-for-running-office-add-ins.md).|


## <a name="outlook-add-in-doesnt-work-correctly"></a>O suplemento do Outlook n?o funciona corretamente

Se um suplemento do Outlook executado no Windows n?o est? funcionando corretamente, tente ativar a depura??o de scripts no Internet Explorer. 


- V? para Ferramentas > **Op??es da Internet** > **Avan?ado**.
    
- Em **Navega??o**, desmarque **Desabilitar depura??o de scripts (Internet Explorer)** e **Desabilitar depura??o de scripts (Outros)**.
    
Recomendamos que voc? desmarque essas configura??es somente para solucionar o problema. Se voc? deixar desmarcado, receber? prompts durante a navega??o. Depois que o problema for resolvido, marque **Desabilitar depura??o de scripts (Internet Explorer)** e **Desabilitar depura??o de scripts (Outros)** novamente.


## <a name="add-in-doesnt-activate-in-office-2013"></a>O suplemento n?o ? ativado no Office 2013

Se o suplemento n?o for ativado quando o usu?rio executar as seguintes etapas:


1. Entrar com a conta da Microsoft no Office 2013.
    
2. Habilitar a verifica??o de duas etapas para a conta da Microsoft.
    
3. Verificar a identidade ao ser solicitado quando tentar inserir um suplemento.
    
Verifique se as atualiza??es mais recentes do Office foram instaladas ou baixe a [atualiza??o do Office 2013](https://support.microsoft.com/en-us/kb/2986156/).


## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>N?o ? poss?vel carregar o suplemento no painel de tarefas ou outros problemas relacionados ao manifesto do suplemento

Confira [Validar e solucionar problemas com seu manifesto](troubleshoot-manifest.md) para depurar problemas do manifesto de suplemento.


## <a name="add-in-dialog-box-cannot-be-displayed"></a>N?o ? poss?vel exibir a caixa de di?logo do suplemento

Quando o usu?rio usa um suplemento do Office, ele ? solicitado a permitir a exibi??o de uma caixa de di?logo. O usu?rio escolhe **Permitir** e, em seguida, recebe a seguinte mensagem de erro:

"As configura??es de seguran?a do navegador nos impedem de criar uma caixa de di?logo. Tente outro navegador ou configure o navegador para que a 'URL' e o dom?nio mostrado na barra de endere?o estejam na mesma zona de seguran?a".

![Captura de tela da mensagem de erro na caixa de di?logo](http://i.imgur.com/3mqmlgE.png)

|**Navegadores afetados**|**Plataformas afetadas**|
|:--------------------|:---------------------|
|Internet Explorer, Microsoft Edge|Office Online|

Para resolver o problema, os administradores ou usu?rios finais podem adicionar o dom?nio do suplemento ? lista de sites confi?veis no Internet Explorer. Use o mesmo procedimento se estiver trabalhando com o navegador Internet Explorer ou Microsoft Edge.

> [!IMPORTANT]
> Caso n?o confie no suplemento, n?o adicione a respectiva URL ? lista de sites confi?veis.

Para adicionar uma URL ? lista de sites confi?veis:

1. No Internet Explorer, escolha o bot?o Ferramentas e v? para **Op??es da Internet** > **Seguran?a**.
2. Escolha a zona de **Sites confi?veis** e escolha **Sites**.
3. Insira a URL exibida na mensagem de erro e escolha **Adicionar**.
4. Tente usar o suplemento novamente. Se o problema persistir, verifique as configura??es de outras zonas de seguran?a e confira se o dom?nio do suplemento est? na mesma zona que a URL exibida na barra de endere?os do aplicativo do Office.

Esse problema ocorre quando a API da caixa de di?logo ? usada no modo pop-up. Para evitar esse problema, use o sinalizador [displayInFrame](https://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync). Isso requer que a p?gina tenha suporte para exibi??o dentro de um iframe. O exemplo a seguir mostra como usar o sinalizador.

```js

Office.context.ui.displayDialogAsync(startAddress, {displayInFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>Altera??es nos comandos de suplemento, incluindo bot?es da faixa de op??es e itens de menu, n?o entram em vigor
?s vezes, as altera??es nos comandos de suplemento, como o ?cone de um bot?o da faixa de op??es ou o texto de um item de menu, n?o parecem entrar em vigor. Limpe o cache do Office das vers?es antigas.

#### <a name="for-windows"></a>No Windows:
Exclua o conte?do da pasta `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.

#### <a name="for-mac"></a>No Mac:
Exclua o conte?do da pasta `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.

#### <a name="for-ios"></a>No iOS:
Chame `window.location.reload(true)` usando o JavaScript no suplemento para for?ar um recarregamento. Outra alternativa ? reinstalar o Office.

## <a name="see-also"></a>Veja tamb?m

- [Depurar suplementos no Office Online](debug-add-ins-in-office-online.md) 
- [Realizar sideload de um suplemento do Office no iPad e no Mac](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Depurar suplementos do Office no iPad e no Mac](debug-office-add-ins-on-ipad-and-mac.md)  
- [Validar e solucionar problemas com seu manifesto](troubleshoot-manifest.md)
    
