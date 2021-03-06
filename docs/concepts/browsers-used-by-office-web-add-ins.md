---
title: Navegadores usados pelos Suplementos do Office
description: Especifica como o sistema operacional e a versão do Office determinam o navegador que é usado pelos suplementos do Office.
ms.date: 02/24/2021
localization_priority: Normal
ms.openlocfilehash: e3297cde10136fad3e044b682957eb6cc60e2e1d
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505210"
---
# <a name="browsers-used-by-office-add-ins"></a>Navegadores usados pelos Suplementos do Office

Os Complementos do Office são aplicativos Web que são exibidos usando iFrames ao executar no Office na Web e usando controles de navegador incorporados no Office para clientes desktop e móveis. Os suplementos também precisam de um mecanismo JavaScript para executar o JavaScript. O navegador incorporado e o mecanismo são fornecidos por um navegador instalado no computador do usuário.

Qual navegador é usado depende do:

- O sistema operacional do computador.
- Se o complemento está sendo executado no Office na Web, no Microsoft 365 ou no Office 2013 ou posterior.

A tabela a seguir mostra qual navegador é usado pelas várias plataformas e sistemas operacionais.

|SISTEMA OPERACIONAL|Versão do Office|WebView2 de borda (baseado em Chromium) instalado?|Navegador|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|qualquer|Office na Web|Não aplicável|O navegador no qual o Office está aberto.|
|Mac|qualquer|Não aplicável|Safari|
|iOS|qualquer|Não aplicável|Safari|
|Android|qualquer|Não aplicável|Chrome|
|Windows 7, 8.1, 10 | não assinatura do Office 2013 ou posterior|Não importa|Internet Explorer 11|
|Windows 7 | Microsoft 365| Não importa | Internet Explorer 11|
|Windows 8.1,<br>Windows 10 ver. &nbsp; < &nbsp; 1903| Microsoft 365 | Não| Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; < &nbsp; 16.0.11629<sup>1</sup>| Não importa|Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.11629 &nbsp; _E_ &nbsp; < &nbsp; 16.0.13530.20424 <sup>1</sup>| Não importa|Microsoft Edge<sup>2, 3</sup> com WebView original (EdgeHTML)|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.13530.20424<sup>1</sup>| Não |Microsoft Edge<sup>2, 3, 4 com</sup> WebView original (EdgeHTML)|
|Windows 8.1<br>Windows 10| Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.13530.20424<sup>1</sup>| Sim<sup>5</sup>|  Microsoft Edge<sup>2, 3, 4</sup> com WebView2 (baseado em Chromium) |

<sup>1</sup> Consulte a página [histórico de atualizações](/officeupdates/update-history-office365-proplus-by-date) e como encontrar [a versão do cliente do Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19) e o canal de atualização para obter mais detalhes.

<sup>2</sup> Quando o Microsoft Edge está sendo usado, o Narrador do Windows 10 (às vezes chamado de "leitor de tela") lê a marca na página que é aberta no `<title>` painel de tarefas. Quando o Internet Explorer 11 está sendo usado, o Narrador lê a barra de título do painel de tarefas, que vem do valor `<DisplayName>` no manifesto de suplemento.

<sup>3</sup> Se o seu complemento incluir o elemento no manifesto, ele usará o Internet Explorer 11 independentemente da versão do Windows ou `Runtimes` do Microsoft 365. Para mais informações, consulte [Runtimes](../reference/manifest/runtimes.md).

<sup>4</sup> A versão do WebView2 para Os Complementos do Office está em andamento. Como resultado, o Microsoft Edge com WebView original (EdgeHTML) ainda pode ser usado para o seu complemento, mesmo quando o computador tem as versões necessárias do Windows e do Office e o controle WebView2 está instalado no computador. Para usuários de canal mensal, esperamos que essa distribuição seja concluída até o final de março de 2021. A distribuição será posteriormente para clientes Semi-Annual Channel. Atualizaremos essa página assim que tiver essas informações disponíveis.

<sup>5</sup> O controle WebView2 inbeddable deve ser instalado além da instalação do Microsoft Edge para que o Office possa in-locar. Para instalá-lo, consulte [Microsoft Edge WebView2 / Embed web content ... com o Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/).




> [!IMPORTANT]
> O Internet Explorer 11 não oferece suporte às versões do JavaScript posteriores a ES5. Se algum dos usuários do seu complemento tiver plataformas que usam o Internet Explorer 11, então para usar a sintaxe e os recursos do ECMAScript 2015 ou posterior, você tem duas opções:
>
> - Escreva seu código no ECMAScript 2015 (também chamado de ES6) ou javaScript posterior ou em TypeScript e compile seu código para JavaScript do ES5 usando um compilador como [o babel](https://babeljs.io/) ou [o tsc](https://www.typescriptlang.org/index.html).
> - Escreva em ECMAScript 2015 ou posterior JavaScript, mas também carregue uma biblioteca de [polifilamento,](https://wikipedia.org/wiki/Polyfill_(programming)) como [core-js,](https://github.com/zloirock/core-js) que permite ao IE executar seu código.
>
> Além disso, o Internet Explorer 11 não oferece suporte a alguns recursos do HTML5, como mídia, gravação e localização.

## <a name="troubleshooting-microsoft-edge-issues"></a>Solução de problemas do Microsoft Edge

### <a name="service-workers-are-not-working"></a>Os Trabalhadores do Serviço não estão funcionando

Os Complementos do Office não suportam Os Funcionários de Serviço quando o [Microsoft Edge WebView](/microsoft-edge/hosting/webview) original é usado. Eles são suportados com o WebView2 de Borda baseado em [Chromium.](/microsoft-edge/hosting/webview2)

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>Barra de rolagem não aparece no painel de tarefas

Por padrão, as barras de rolagem no Microsoft Edge estão ocultas até que você tenha passado. Para garantir que a barra de rolagem fique sempre visível, o estilo de CSS que se aplica ao elemento `<body>` das páginas no painel de tarefas deve incluir a propriedade [(-ms- reoverflow-style)](https://developer.mozilla.org/docs/Archive/Web/CSS/-ms-overflow-style) e deve ser definida como `scrollbar`.

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>Ao depurar com o Microsoft Edge DevTools, o suplemento falha ou recarrega

A definição de pontos de interrupção nas [DevTools do Microsoft Edge](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) pode fazer o Office pensar que o suplemento está travado. Ele recarrega automaticamente o suplemento quando isso acontece. Para evitar isso, adicione a seguinte chave do registro e valor ao computador de desenvolvimento `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`:.

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>Quando o suplemento tentar abrir, o erro “ADD-IN ERROR não é possível abrir este suplemento a partir do localhost" acontece

Uma causa conhecida é que o Microsoft Edge exige que o localhost tenha uma isenção de auto-retorno no computador de desenvolvimento. Siga as instruções em [não é possível abrir o suplemento do localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).

### <a name="get-errors-trying-to-download-a-pdf-file"></a>Obter erros ao tentar baixar um arquivo PDF

Não há suporte para download direto de blobs como arquivos PDF em um complemento quando o Edge é o navegador. A solução alternativa é criar um aplicativo Web simples que baixa blobs como arquivos PDF. No seu complemento, chame o método `Office.context.ui.openBrowserWindow(url)` e passe a URL do aplicativo Web. Isso abrirá o aplicativo Web em uma janela do navegador fora do Office.

## <a name="see-also"></a>Confira também

- [Requisitos para a Execução de Suplementos do Office](requirements-for-running-office-add-ins.md)
