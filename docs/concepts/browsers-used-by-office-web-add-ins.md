---
title: Navegadores usados pelos Suplementos do Office
description: Especifica como o sistema operacional e a versão do Office determinam o navegador que é usado pelos suplementos do Office.
ms.date: 08/13/2020
localization_priority: Normal
ms.openlocfilehash: 544388014bfef0dd647a79d655a173d09f5a4ff7
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408436"
---
# <a name="browsers-used-by-office-add-ins"></a>Navegadores usados pelos Suplementos do Office

Os suplementos do Office são aplicativos Web que são exibidos usando iFrames ao executar no Office na Web e usando controles de navegador incorporados no Office para clientes móveis e de área de trabalho. Os suplementos também precisam de um mecanismo JavaScript para executar o JavaScript. O navegador incorporado e o mecanismo são fornecidos por um navegador instalado no computador do usuário.

Qual navegador é usado depende do:

- Sistema operacional do computador.
- Se o suplemento está sendo executado no Office na Web, no Microsoft 365 ou no Office 2013 ou posterior que não está em assinatura.

A tabela a seguir mostra qual navegador é usado pelas várias plataformas e sistemas operacionais.

|Opera|Versão do Office|O Edge WebView2 (baseado em Chromium) está instalado?|Navegador|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|qualquer|Office na Web|Não aplicável|O navegador no qual o Office está aberto.|
|Mac|qualquer|Não aplicável|Safari|
|iOS|qualquer|Não aplicável|Safari|
|Android|qualquer|Não aplicável|Chrome|
|Windows 7, 8,1, 10 | Office 2013 não inscrito ou posterior|Não importa|Internet Explorer 11|
|Windows 7 | Microsoft 365| Não importa | Internet Explorer 11|
|Windows 8,1,<br>Windows 10 ver. &nbsp; < &nbsp; 1903| Microsoft 365 | Não| Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; < &nbsp; 16.0.11629<sup>1</sup>| Não importa|Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.11629 &nbsp; _e_ &nbsp; < &nbsp; 16.0.13127.20082<sup>1</sup>| Não importa|Microsoft Edge<sup>2, 3</sup> com WebView original (EdgeHTML)|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.13127.20082<sup>1</sup>| Não |Microsoft Edge<sup>2, 3</sup> com WebView original (EdgeHTML)|
|Windows 8.1<br>Windows 10| Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.13127.20082<sup>1</sup>| Sim|  Consulte a observação 4 abaixo. |

<sup>1</sup> consulte a [página Histórico de atualizações](/officeupdates/update-history-office365-proplus-by-date) e como [encontrar sua versão e canal de atualização do cliente Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19) para obter mais detalhes.

<sup>2</sup> quando o Microsoft Edge está sendo usado, o Windows 10 Narrator (às vezes chamado de "leitor de tela") lê a `<title>` marca na página que é aberta no painel de tarefas. Quando o Internet Explorer 11 está sendo usado, o Narrador lê a barra de título do painel de tarefas, que vem do valor `<DisplayName>` no manifesto de suplemento.

<sup>3</sup> se o suplemento incluir o `Runtimes` elemento no manifesto, ele usará o Internet Explorer 11 independentemente da versão do Windows ou do Microsoft 365. Para mais informações, consulte [Runtimes](../reference/manifest/runtimes.md).

<sup>4</sup> o navegador usado para essa combinação de versões depende do canal de atualização da assinatura do Microsoft 365. Se o usuário estiver no [canal beta](https://insider.office.com/join/windows) (antigo canal de insider), o Office usa o Microsoft Edge com o WebView2 (baseado em Chromium). Para qualquer outro canal, o Office usa o Microsoft Edge com o WebView original (EdgeHTML). O suporte para WebView2 em outros canais é esperado no início de 2021.

> [!IMPORTANT]
> O Internet Explorer 11 não oferece suporte às versões do JavaScript posteriores a ES5. Se qualquer um dos usuários do seu suplemento tiver plataformas que usam o Internet Explorer 11, use a sintaxe e os recursos do ECMAScript 2015 ou posterior, você tem duas opções:
>
> - Escreva seu código no ECMAScript 2015 (também chamado de ES6) ou em JavaScript posterior, ou em TypeScript, e compile seu código para ES5 JavaScript usando um compilador como [Babel](https://babeljs.io/) ou [TSC](https://www.typescriptlang.org/index.html).
> - Escreva em JavaScript 2015 ou superior, mas também carregue uma biblioteca de [polipreenchimento](https://wikipedia.org/wiki/Polyfill_(programming)) , como [Core-js](https://github.com/zloirock/core-js) , que permite ao ie executar o código.
>
> Além disso, o Internet Explorer 11 não oferece suporte a alguns recursos do HTML5, como mídia, gravação e localização.

## <a name="troubleshooting-microsoft-edge-issues"></a>Solucionando problemas do Microsoft Edge

### <a name="service-workers-are-not-working"></a>Os funcionários de serviço não estão funcionando

Os suplementos do Office não dão suporte a trabalhadores de serviço quando o [Microsoft Edge WebView](/microsoft-edge/hosting/webview) original é usado. Eles são compatíveis com o [WebView2 de borda baseado em Chromium](/microsoft-edge/hosting/webview2).

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
