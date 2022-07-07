---
title: Navegadores usados pelos Suplementos do Office
description: Especifica como o sistema operacional e a versão do Office determinam o navegador que é usado pelos suplementos do Office.
ms.date: 05/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1fedeb7f7e1e972a2a7fe4befa5a990ff8cc698d
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659651"
---
# <a name="browsers-used-by-office-add-ins"></a>Navegadores usados pelos Suplementos do Office

Os Suplementos do Office são aplicativos Web exibidos usando iFrames durante a execução Office na Web. No Office para clientes desktop e móveis, os Suplementos do Office usam um controle de navegador inserido (também conhecido como modo de exibição da Web). Os suplementos também precisam de um mecanismo JavaScript para executar o JavaScript. O navegador inserido e o mecanismo são fornecidos por um navegador instalado no computador do usuário.

Qual navegador é usado depende do:

- O sistema operacional do computador.
- Se o suplemento está em execução no Office na Web, No Microsoft 365 ou no Office 2013 ou posterior.

> [!IMPORTANT]
> **Internet Explorer ainda usado em Suplementos do Office**
>
> Algumas combinações de plataformas e versões do Office, incluindo versões de compra única por meio do Office 2019, ainda usam o controle webview que vem com o Internet Explorer 11 para hospedar suplementos, conforme explicado neste artigo. Recomendamos (mas não exige) que você continue a dar suporte a essas combinações, pelo menos de maneira mínima, fornecendo aos usuários do seu suplemento uma mensagem de falha normal quando o suplemento é iniciado no modo de exibição da Web do Internet Explorer. Lembre-se destes pontos adicionais:
>
> - Office na Web abre mais no Internet Explorer. Consequentemente, [o AppSource](/office/dev/store/submit-to-appsource-via-partner-center) não testa mais suplementos no Office na Web usando o Internet Explorer como navegador.
> - O AppSource ainda testa combinações de versões da plataforma e da área de trabalho *do Office que* usam o Internet Explorer, no entanto, ele só emite um aviso quando o suplemento não dá suporte ao Internet Explorer; o suplemento não é rejeitado pelo AppSource.
> - A [Script Lab não dá](../overview/explore-with-script-lab.md) mais suporte ao Internet Explorer.
>
> Para obter mais informações sobre como dar suporte ao Internet Explorer e configurar uma mensagem de falha normal em seu suplemento, consulte [Suporte do Internet Explorer 11](../develop/support-ie-11.md).

A tabela a seguir mostra qual navegador é usado pelas várias plataformas e sistemas operacionais.

|SO|Versão do Office|Edge WebView2 (baseado Chromium) instalado?|Navegador|
|:-----|:-----|:-----|:-----|
|qualquer|Office na Web|Não aplicável|O navegador no qual o Office está aberto.<br>(Mas observe que Office na Web não será aberto no Internet Explorer.<br>A tentativa de fazer isso abre Office na Web no Edge.) |
|Mac|qualquer|Não aplicável|Safari com WKWebView|
|iOS|qualquer|Não aplicável|Safari com WKWebView|
|Android|qualquer|Não aplicável|Chrome|
|Windows 7, 8.1, 10, 11 | Office 2013 sem assinatura para Office 2019|Não importa, não importa.|Internet Explorer 11|
|Windows 10, 11 | não assinatura Office 2021 ou posterior|Sim|Microsoft Edge<sup>1</sup> com WebView2 (Chromium baseado)|
|Windows 7 | Microsoft 365| Não importa, não importa. | Internet Explorer 11|
|Windows 8.1,<br>Windows 10 ver.&nbsp;<&nbsp; 1903| Microsoft 365 | Não| Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp; 1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;<&nbsp; 16.0.11629<sup>2</sup>| Não importa, não importa.|Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp; 1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;>=&nbsp; 16.0.11629&nbsp;_E_&nbsp;<&nbsp;16.0.13530.20424 <sup>2</sup>| Não importa, não importa.|Microsoft Edge<sup>1, 3</sup> com WebView original (EdgeHTML)|
|Windows 10 ver.&nbsp;>=&nbsp; 1903,<br>Janela 11 | Microsoft 365 ver.&nbsp;>=&nbsp; 16.0.13530.20424<sup>2</sup>| Não |Microsoft Edge<sup>1, 3</sup> com WebView original (EdgeHTML)|
|Windows 8.1<br>Windows 10,<br>Windows 11| Microsoft 365 ver.&nbsp;>=&nbsp; 16.0.13530.20424<sup>2</sup>| Sim<sup>4</sup>|  Microsoft Edge<sup>1</sup> com WebView2 (Chromium baseado) |

<sup>1</sup> Quando o Microsoft Edge está sendo usado, o Narrador do Windows (às vezes chamado de "leitor de tela") `<title>` lê a marca na página que é aberta no painel de tarefas. Quando o Internet Explorer 11 está sendo usado, o Narrador lê a barra de título do painel de tarefas, **\<DisplayName\>** que vem do valor no manifesto do suplemento.

<sup>2</sup> Consulte a página [histórico de atualizações](/officeupdates/update-history-office365-proplus-by-date) e como encontrar [a versão do cliente do Office e o canal de atualização](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19) para obter mais detalhes.

<sup>3</sup> Se o suplemento **\<Runtimes\>** incluir o elemento no manifesto, ele não usará o Microsoft Edge com o WebView original (EdgeHTML). Se as condições para usar o Microsoft Edge com WebView2 (Chromium baseadas em Chromium) forem atendidas, o suplemento usará esse navegador. Caso contrário, ele usará o Internet Explorer 11, independentemente da versão do Windows ou do Microsoft 365. Para mais informações, consulte [Runtimes](/javascript/api/manifest/runtimes).

<sup>4</sup> Em versões do Windows anteriores Windows 11, o controle WebView2 deve ser instalado para que o Office possa inseri-lo. Ele é instalado com o Microsoft 365, versão 2101 ou posterior e com compra única Office 2021 ou posterior; mas não é instalado automaticamente com o Microsoft Edge. Se você tiver uma versão anterior do Microsoft 365 ou do Office de compra única, use as instruções para instalar o controle no [Microsoft Edge WebView2/Inserir conteúdo da Web... com o Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/). Em builds do Microsoft 365 anteriores a 16.0.14326.xxxxx, você também deve criar a chave do **RegistroHKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Win32WebView2** e definir seu valor como `dword:00000001`.

> [!IMPORTANT]
> O Internet Explorer 11 não oferece suporte às versões do JavaScript posteriores a ES5. Se algum dos usuários do suplemento tiver plataformas que usam o Internet Explorer 11, para usar a sintaxe e os recursos do ECMAScript 2015 ou posterior, você terá duas opções.
>
> - Escreva seu código no ECMAScript 2015 (também chamado de ES6) ou em JavaScript posterior ou em TypeScript e, em seguida, compile seu código em JavaScript ES5 usando um compilador como [babel](https://babeljs.io/) ou [tsc](https://www.typescriptlang.org/index.html).
> - Escreva no ECMAScript 2015 ou em JavaScript posterior, mas também carregue uma biblioteca de [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) , como [core-js](https://github.com/zloirock/core-js) , que permite que o IE execute seu código.
>
> Para obter mais informações sobre essas opções, consulte [Suporte do Internet Explorer 11](../develop/support-ie-11.md).
>
> Além disso, o Internet Explorer 11 não oferece suporte a alguns recursos do HTML5, como mídia, gravação e localização. Para saber mais, confira [Determinar em runtime se o suplemento está em execução no Internet Explorer](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer).

## <a name="troubleshooting-microsoft-edge-issues"></a>Solução de problemas do Microsoft Edge

### <a name="service-workers-are-not-working"></a>Os Trabalhadores do Serviço não estão funcionando

Os Suplementos do Office não dão suporte a Service Workers quando o Microsoft Edge WebView original, [EdgeHTML](https://en.wikipedia.org/wiki/EdgeHTML), é usado. Eles têm suporte com o [Edge WebView2 Chromium baseado em Chromium](/microsoft-edge/hosting/webview2).

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>Barra de rolagem não aparece no painel de tarefas

Por padrão, as barras de rolagem no Microsoft Edge estão ocultas até que você tenha passado. Para garantir que a barra de rolagem fique sempre visível, o estilo de CSS que se aplica ao elemento `<body>` das páginas no painel de tarefas deve incluir a propriedade [(-ms- reoverflow-style)](https://developer.mozilla.org/docs/Web/CSS/Microsoft_Extensions) e deve ser definida como `scrollbar`.

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>Ao depurar com o Microsoft Edge DevTools, o suplemento falha ou recarrega

A definição de pontos de interrupção nas [DevTools do Microsoft Edge](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) pode fazer o Office pensar que o suplemento está travado. Ele recarrega automaticamente o suplemento quando isso acontece. Para evitar isso, adicione a seguinte chave do registro e valor ao computador de desenvolvimento `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`:.

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>Quando o suplemento tentar abrir, o erro “ADD-IN ERROR não é possível abrir este suplemento a partir do localhost" acontece

Uma causa conhecida é que o Microsoft Edge exige que o localhost tenha uma isenção de auto-retorno no computador de desenvolvimento. Siga as instruções em [não é possível abrir o suplemento do localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).

### <a name="get-errors-trying-to-download-a-pdf-file"></a>Obter erros ao tentar baixar um arquivo PDF

Não há suporte para o download direto de blobs como arquivos PDF em um suplemento quando o Edge é o navegador. A solução alternativa é criar um aplicativo Web simples que baixa blobs como arquivos PDF. No suplemento, chame o método `Office.context.ui.openBrowserWindow(url)` e passe a URL do aplicativo Web. Isso abrirá o aplicativo Web em uma janela do navegador fora do Office.

## <a name="see-also"></a>Confira também

- [Requisitos para a Execução de Suplementos do Office](requirements-for-running-office-add-ins.md)
