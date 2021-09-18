---
title: Navegadores usados pelos Suplementos do Office
description: Especifica como o sistema operacional e a versão do Office determinam o navegador que é usado pelos suplementos do Office.
ms.date: 09/10/2021
ms.localizationpriority: medium
ms.openlocfilehash: 77cf0b6888100eee6fa6d90f221dc680a9991a7e
ms.sourcegitcommit: 3fe9e06a52c57532e7968dc007726f448069f48d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/18/2021
ms.locfileid: "59443521"
---
# <a name="browsers-used-by-office-add-ins"></a>Navegadores usados pelos Suplementos do Office

Office Os complementos são aplicativos Web que são exibidos usando iFrames ao executar no Office na Web e usando controles de navegador incorporados no Office para clientes desktop e móveis. Os suplementos também precisam de um mecanismo JavaScript para executar o JavaScript. O navegador incorporado e o mecanismo são fornecidos por um navegador instalado no computador do usuário.

Qual navegador é usado depende do:

- O sistema operacional do computador.
- Se o add-in está sendo executado em Office na Web, Microsoft 365 ou não de assinatura Office 2013 ou posterior.

> [!IMPORTANT]
> **Internet Explorer ainda usado em Office de complementos**
>
> A Microsoft está encerrando o suporte para o Internet Explorer, mas isso não afeta significativamente Office Desempios. Algumas combinações de plataformas e versões Office, incluindo todas as versões de compra única por meio do Office 2019, continuarão a usar o controle webview que vem com o Internet Explorer 11 para hospedar os complementos, conforme explicado neste artigo. Além disso, o suporte a essas combinações e, portanto, para o Internet Explorer, ainda é necessário para os complementos enviados ao [AppSource](/office/dev/store/submit-to-appsource-via-partner-center). Duas coisas *estão mudando:*
>
> - O AppSource não testa mais os Office na Web usando o Internet Explorer como navegador. Mas o AppSource ainda testa combinações de plataforma e Office *desktop* que usam o Internet Explorer.
> - A [Script Lab não](../overview/explore-with-script-lab.md) dá mais suporte ao Internet Explorer.

A tabela a seguir mostra qual navegador é usado pelas várias plataformas e sistemas operacionais.

|SISTEMA OPERACIONAL|Versão do Office|WebView2 de borda (Chromium baseado em dados) instalado?|Navegador|
|:-----|:-----|:-----|:-----|
|qualquer|Office na Web|Não aplicável|O navegador no qual o Office está aberto.|
|Mac|qualquer|Não aplicável|Safari|
|iOS|qualquer|Não aplicável|Safari|
|Android|qualquer|Não aplicável|Chrome|
|Windows 7, 8.1, 10 | non-subscription Office 2013 to Office 2019|Não importa|Internet Explorer 11|
|Windows 10 | non-subscription Office 2021 or later|Sim|Microsoft Edge<sup>1</sup> com WebView2 (Chromium baseado em Chromium)|
|Windows 7 | Microsoft 365| Não importa | Internet Explorer 11|
|Windows 8.1,<br>Windows 10 ver. &nbsp; < &nbsp; 1903| Microsoft 365 | Não| Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; < &nbsp; 16.0.11629<sup>2</sup>| Não importa|Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.11629 &nbsp; _E_ &nbsp; < &nbsp; 16.0.13530.20424 <sup>2</sup>| Não importa|Microsoft Edge<sup>1, 3</sup> com WebView original (EdgeHTML)|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.13530.20424<sup>2</sup>| Não |Microsoft Edge<sup>1, 3</sup> com WebView original (EdgeHTML)|
|Windows 8.1<br>Windows 10| Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.13530.20424<sup>2</sup>| Sim<sup>4</sup>|  Microsoft Edge<sup>1</sup> com WebView2 (Chromium baseado em Chromium) |

<sup>1</sup> Quando o Microsoft Edge está sendo usado, o narrador Windows 10 (às vezes chamado de "leitor de tela") lê a marca na página que é aberta no `<title>` painel de tarefas. Quando o Internet Explorer 11 está sendo usado, o Narrador lê a barra de título do painel de tarefas, que vem do valor `<DisplayName>` no manifesto de suplemento.

<sup>2</sup> Consulte a página [histórico de](/officeupdates/update-history-office365-proplus-by-date) atualizações e como encontrar sua versão Office cliente e [o](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19) canal de atualização para obter mais detalhes.

<sup>3</sup> Se o seu complemento incluir o elemento no manifesto, ele não usará Microsoft Edge com o `<Runtimes>` WebView original (EdgeHTML). Se as condições de uso Microsoft Edge webView2 (Chromium baseadas em Chromium) são atendidas, o complemento usa esse navegador. Caso contrário, ele usa o Internet Explorer 11, independentemente da Windows ou Microsoft 365 versão. Para mais informações, consulte [Runtimes](../reference/manifest/runtimes.md).

<sup>4 O</sup> controle WebView2 inbeddable deve ser instalado para que Office possa in-lo, e ele não é instalado automaticamente com o Edge. Ele é instalado com Microsoft 365, versão 2101 ou posterior. Se você tiver uma versão anterior do Microsoft 365, use as instruções para instalar o controle em [Microsoft Edge WebView2 / Incorporar conteúdo da Web... com Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/).

> [!IMPORTANT]
> O Internet Explorer 11 não oferece suporte às versões do JavaScript posteriores a ES5. Se algum dos usuários do seu complemento tiver plataformas que usam o Internet Explorer 11, para usar a sintaxe e os recursos do ECMAScript 2015 ou posterior, você terá duas opções.
>
> - Escreva seu código no ECMAScript 2015 (também chamado de ES6) ou javaScript posterior ou em TypeScript e compile seu código para JavaScript do ES5 usando um compilador como [o babel](https://babeljs.io/) ou [o tsc](https://www.typescriptlang.org/index.html).
> - Escreva em ECMAScript 2015 ou posterior JavaScript, mas também carregue uma biblioteca de [polifilamento,](https://en.wikipedia.org/wiki/Polyfill_(programming)) como [core-js,](https://github.com/zloirock/core-js) que permite ao IE executar seu código.
>
> Para obter mais informações sobre essas opções, consulte [Support Internet Explorer 11](../develop/support-ie-11.md).
>
> Além disso, o Internet Explorer 11 não oferece suporte a alguns recursos do HTML5, como mídia, gravação e localização.

## <a name="troubleshooting-microsoft-edge-issues"></a>Solução de Microsoft Edge problemas

### <a name="service-workers-are-not-working"></a>Os Trabalhadores do Serviço não estão funcionando

Office Os complementos não suportam Os Trabalhadores do Serviço quando o webView original Microsoft Edge WebView, [EdgeHTML](https://en.wikipedia.org/wiki/EdgeHTML), é usado. Eles são suportados com o [Chromium WebView2](/microsoft-edge/hosting/webview2)baseado em Borda.

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>Barra de rolagem não aparece no painel de tarefas

Por padrão, as barras de rolagem no Microsoft Edge estão ocultas até que você tenha passado. Para garantir que a barra de rolagem fique sempre visível, o estilo de CSS que se aplica ao elemento `<body>` das páginas no painel de tarefas deve incluir a propriedade [(-ms- reoverflow-style)](https://developer.mozilla.org/docs/Web/CSS/Microsoft_Extensions) e deve ser definida como `scrollbar`.

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>Ao depurar com o Microsoft Edge DevTools, o suplemento falha ou recarrega

A definição de pontos de interrupção nas [DevTools do Microsoft Edge](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) pode fazer o Office pensar que o suplemento está travado. Ele recarrega automaticamente o suplemento quando isso acontece. Para evitar isso, adicione a seguinte chave do registro e valor ao computador de desenvolvimento `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`:.

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>Quando o suplemento tentar abrir, o erro “ADD-IN ERROR não é possível abrir este suplemento a partir do localhost" acontece

Uma causa conhecida é que o Microsoft Edge exige que o localhost tenha uma isenção de auto-retorno no computador de desenvolvimento. Siga as instruções em [não é possível abrir o suplemento do localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).

### <a name="get-errors-trying-to-download-a-pdf-file"></a>Obter erros ao tentar baixar um arquivo PDF

Não há suporte para download direto de blobs como arquivos PDF em um complemento quando o Edge é o navegador. A solução alternativa é criar um aplicativo Web simples que baixa blobs como arquivos PDF. No seu complemento, chame o método `Office.context.ui.openBrowserWindow(url)` e passe a URL do aplicativo Web. Isso abrirá o aplicativo Web em uma janela do navegador fora do Office.

## <a name="see-also"></a>Confira também

- [Requisitos para a Execução de Suplementos do Office](requirements-for-running-office-add-ins.md)
