---
title: Depurar suplementos do Office
description: Encontre a diretrizes de depuração do Suplemento do Office para seu ambiente de desenvolvimento
ms.date: 02/17/2022
ms.localizationpriority: high
ms.openlocfilehash: 67afa308625c4f18775c3470c2bc2410af37905c
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340376"
---
# <a name="overview-of-debugging-office-add-ins"></a>Visão geral da depuração de Suplementos do Office

A depuração de Suplementos do Office é essencialmente a mesma que a depuração qualquer aplicativo Web. No entanto, um único conjunto de ferramentas não funcionará para todos os desenvolvedores de suplementos. Isso ocorre porque os suplementos podem ser desenvolvidos em diferentes sistemas operacionais e executados em várias plataformas. Este artigo ajuda você a encontrar as diretrizes de depuração detalhadas para seu ambiente de desenvolvimento.

> [!TIP]
> Este artigo está preocupado com a depuração no sentido estrito de definir pontos de interrupção e percorrer o código. Para obter as diretrizes sobre testes e solução de problemas, comece com [Testar Suplementos do Office](test-debug-office-add-ins.md) e [Solução de problemas de erros de desenvolvimento com Suplementos do Office](troubleshoot-development-errors.md).

> [!NOTE]
> Embora você deva *testar* seu suplemento em todas as plataformas às quais deseja oferecer suporte, você raramente precisará *depurar* em um ambiente diferente do seu computador de desenvolvimento. Por esse motivo, este artigo utiliza “seu computador de desenvolvimento” e “seu ambiente de desenvolvimento” para se referir ao ambiente no qual você está depurando. Se um problema no código ocorrer apenas em uma plataforma diferente daquela em seu computador de desenvolvimento e você precisar definir pontos de interrupção ou percorrer o código para resolvê-lo, o ambiente no qual você está depurando não é literalmente seu ambiente de desenvolvimento.

## <a name="server-side-or-client-side"></a>Do lado do servidor ou do lado do cliente?

Depurar o código do lado do servidor de um suplemento do Office é o mesmo que depurar o lado do servidor de qualquer aplicativo Web. Veja as instruções de depuração do seu IDE ou de outras ferramentas. A seguir estão alguns exemplos de algumas das ferramentas mais populares.

- [Depurar aplicativos ASP.NET ou ASP.NET Core no Visual Studio](/visualstudio/debugger/how-to-enable-debugging-for-aspnet-applications)
- [Depuração Expressa](https://expressjs.com/en/guide/debugging.html)
- [Guia de depuração do Node.js](https://nodejs.org/en/docs/guides/debugging-getting-started/)
- [Depuração do Node.js no VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging)
- [Depuração do Webpack](https://webpack.js.org/contribute/debugging/)

O restante deste artigo está preocupado apenas com a depuração do JavaScript do lado do cliente (que pode ser transpilado do TypeScript).

Para encontrar as diretrizes para depurar o código do lado do cliente, a primeira variável é o sistema operacional do seu computador de desenvolvimento.

- [Windows](#debug-on-windows)
- [Mac](#debug-on-mac)
- [Linux ou outra variante Unix](#debug-on-linux)

## <a name="debug-on-windows"></a>Depurar no Windows

A seguir, as diretrizes gerais para a depuração no Windows. Há instruções especiais para a depuração de funções personalizadas sem interface do usuário no Excel e suplementos baseados em eventos no Outlook. Consulte [Casos especiais no Windows](#special-cases-in-windows) posteriormente nesta seção. A depuração no Windows depende do seu IDE:

- **Visual Studio**: depurar usando as ferramentas F12 do navegador. Consulte [Depurar Suplementos do Office no Visual Studio](../develop/debug-office-add-ins-in-visual-studio.md).
- **Visual Studio Code**: Depure usando a [Extensão do Depurador de Suplemento para Visual Studio Code](debug-with-vs-extension.md).
- **Qualquer outro IDE** (ou você não quer depurar dentro do seu IDE): use as ferramentas de desenvolvedor associadas ao runtime do navegador que os suplementos usam no seu computador de desenvolvimento. Consulte uma das seguintes opções:

    - [Depurar os suplementos usando as ferramentas de desenvolvedor para o Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
    - [Depurar suplementos usando ferramentas de desenvolvedor para Edge Legacy](debug-add-ins-using-devtools-edge-legacy.md)
    - [Depurar suplementos usando ferramentas de desenvolvedor no Microsoft Edge (baseado em Chromium)](debug-add-ins-using-devtools-edge-chromium.md)

Para obter informações sobre qual runtime do navegador está sendo usado, confira [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

### <a name="special-cases-in-windows"></a>Casos especiais no Windows

Para depurar funções personalizadas sem interface do usuário no Windows, confira [Depuração de funções personalizadas sem interface do usuário](../excel/custom-functions-debugging.md).

Para depurar suplementos baseados em eventos no Outlook, confira [Depurar seu suplemento do Outlook baseado em eventos](../outlook/debug-autolaunch.md). O processo exige o Visual Studio Code.

## <a name="debug-on-mac"></a>Depurar no Mac

Veja a seguir diretrizes gerais para depuração no Mac. Existem instruções especiais para depurar funções personalizadas sem interface do usuário no Excel. Consulte [Casos especiais no Mac](#special-cases-in-mac) posteriormente nesta seção.

- Se você estiver usando o Visual Studio Code, depure usando a [Extensão do Depurador de Suplemento para Visual Studio Code ](debug-with-vs-extension.md).
- Para qualquer outro IDE, use o Safari Web Inspector. As instruções estão em [Depurar Suplementos do Office em um Mac](debug-office-add-ins-on-ipad-and-mac.md).

### <a name="special-cases-in-mac"></a>Casos especiais no Mac

Para depurar funções personalizadas sem interface do usuário no Mac, consulte [Depuração de funções personalizadas sem interface do usuário](../excel/custom-functions-debugging.md).

## <a name="debug-on-linux"></a>Depurar no Linux

Não há versões da área de trabalho do Office para Linux, portanto, você precisará [realizar o sideload do suplemento do Office na Web](sideload-office-add-ins-for-testing.md) para testá-lo e depurá-lo. As diretrizes de depuração estão nos [Suplementos de depuração no Office na Web](debug-add-ins-in-office-online.md).

> [!NOTE]
> Não recomendamos que você desenvolva Suplementos do Office em um computador Linux, exceto no caso incomum em que você pode ter certeza de que todos os usuários do suplemento acessarão o suplemento por meio do Office na Web a partir de um computador Linux.

## <a name="debug-add-ins-in-staging-or-production"></a>Depurar suplementos em preparo ou produção

Para depurar um suplemento que já está em preparo ou produção, anexe um depurador da interface do usuário do suplemento. Para obter instruções, [Anexe um depurador no painel de tarefas](attach-debugger-from-task-pane.md).