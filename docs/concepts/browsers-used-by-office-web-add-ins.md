---
title: Navegadores usados pelos Suplementos do Office
description: Especifica como o sistema operacional e a versão do Office determinam o navegador que é usado pelos suplementos do Office.
ms.date: 04/21/2020
localization_priority: Normal
ms.openlocfilehash: 9ef4b6d4c09140fc6d6bb04eca51d845b79b6dc7
ms.sourcegitcommit: 3355c6bd64ecb45cea4c0d319053397f11bc9834
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/22/2020
ms.locfileid: "43744849"
---
# <a name="browsers-used-by-office-add-ins"></a>Navegadores usados pelos Suplementos do Office

Os suplementos do Office são aplicativos Web exibidos usando iFrames durante a execução do Office na Web e no uso de controles de navegador incorporados no Office para clientes desktops e móveis. Os suplementos também precisam de um mecanismo JavaScript para executar o JavaScript. O navegador incorporado e o mecanismo são fornecidos por um navegador instalado no computador do usuário.

Qual navegador é usado depende do:

- Sistema operacional do computador.
- Se o suplemento está em execução no Office na Web, no Office 365 ou no Office 2013 sem assinatura ou posterior.

A tabela a seguir mostra qual navegador é usado pelas várias plataformas e sistemas operacionais.

|**SO / Plataforma**|**Navegador**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|Office na Web|O navegador no qual o Office está aberto.|
|Mac|Safari|
|iOS|Safari|
|Android|Chrome|
|Windows / Office 2013 sem assinatura ou posterior.|Internet Explorer 11|
|Versão do Windows 10 < 1903 / Office 365|Internet Explorer 11|
|Versão do Windows 10 >= 1903/Office 365 ver < 16.0.11629<sup>1</sup>|Internet Explorer 11|
|Versão do Windows 10 >= 1903/Office 365 ver >= 16.0.11629<sup>1</sup>|Microsoft Edge<sup>2</sup>|

<sup>1</sup> consulte a [página Histórico de atualizações](/officeupdates/update-history-office365-proplus-by-date) e como [encontrar sua versão e canal de atualização do cliente Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19) para obter mais detalhes.

<sup>2</sup> quando o Microsoft Edge está sendo usado, o Windows 10 Narrator (às vezes chamado de "leitor de tela" `<title>` ) lê a marca na página que é aberta no painel de tarefas. Quando o Internet Explorer 11 está sendo usado, o Narrador lê a barra de título do painel de tarefas, que vem do valor `<DisplayName>` no manifesto de suplemento.

> [!IMPORTANT]
> O Internet Explorer 11 não oferece suporte às versões do JavaScript posteriores a ES5. Se qualquer um dos usuários de suplemento tiverem plataformas com Internet Explorer 11, para que seja possível usar a sintaxe e os recursos do ECMAScript 2015 ou posterior, você precisará fazer o transpile do seu JavaScript para o ES5 ou usar um polyfill. Além disso, o Internet Explorer 11 não oferece suporte a alguns recursos do HTML5, como mídia, gravação e localização.

## <a name="troubleshooting-microsoft-edge-issues"></a>Solucionar problemas do Microsoft Edge

### <a name="service-workers-are-not-working"></a>Os funcionários de serviço não estão funcionando

Os suplementos do Office não dão suporte a trabalhadores de serviço no [Microsoft Edge WebView](/microsoft-edge/hosting/webview). Confira a [visão geral dos suplementos do Office](../overview/office-add-ins.md) para obter os recursos mais recentes suportados no controle de WebView de borda. Estamos trabalhando duro para trazer o novo [WebView2 de borda baseado em Chromium](/microsoft-edge/hosting/webview2) para a plataforma de suplementos do Office, que esperamos que seja compatível com os funcionários de serviço.

### <a name="chromium-based-edge-is-installed-on-my-development-computer-but-my-add-in-does-not-use-it"></a>A borda baseada em Chromium está instalada no meu computador de desenvolvimento, mas meu suplemento não a usa

O navegador base no [Microsoft Edge](https://support.microsoft.com/help/4501095/download-the-new-microsoft-edge-based-on-chromium) foi alterado para Chromium. A base mais antiga, chamada EdgeHTML, não é removida quando a borda baseada no Chromium é instalada. O Office ainda usará a base de EdgeHTML para suplementos até que uma compilação do Office 365 que suporte o Chromium esteja instalada no computador. Esperamos que essas compilações sejam fornecidas no 2020. Eles provavelmente aparecerão no canal de insiders na primeira metade do ano.

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>Barra de rolagem não aparece no painel de tarefas

Por padrão, as barras de rolagem no Microsoft Edge estão ocultas até que você tenha passado. Para garantir que a barra de rolagem fique sempre visível, o estilo de CSS que se aplica ao elemento `<body>` das páginas no painel de tarefas deve incluir a propriedade [(-ms- reoverflow-style)](https://developer.mozilla.org/docs/Web/CSS/-ms-overflow-style) e deve ser definida como `scrollbar`. 

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>Ao depurar com o Microsoft Edge DevTools, o suplemento falha ou recarrega

A definição de pontos de interrupção nas [DevTools do Microsoft Edge](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) pode fazer o Office pensar que o suplemento está travado. Ele recarrega automaticamente o suplemento quando isso acontece. Para evitar isso, adicione a seguinte chave do registro e valor ao computador de desenvolvimento `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`:.

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>Quando o suplemento tentar abrir, o erro “ADD-IN ERROR não é possível abrir este suplemento a partir do localhost" acontece

Uma causa conhecida é que o Microsoft Edge exige que o localhost tenha uma isenção de auto-retorno no computador de desenvolvimento. Siga as instruções em [não é possível abrir o suplemento do localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).

### <a name="get-errors-trying-to-download-a-pdf-file"></a>Obter erros ao tentar baixar um arquivo PDF

O download de BLOBs diretamente como arquivos PDF em um suplemento não é suportado quando Edge é o navegador. A solução alternativa é criar um aplicativo Web simples que baixe BLOBs como arquivos PDF. No seu suplemento, chame o `Office.context.ui.openBrowserWindow(url)` método e passe a URL do aplicativo Web. Isso abrirá o aplicativo Web em uma janela do navegador fora do Office.

## <a name="see-also"></a>Confira também

- [Requisitos para a Execução de Suplementos do Office](requirements-for-running-office-add-ins.md)
