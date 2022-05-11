---
title: Teste do Internet Explorer 11
description: Teste seu Office suplemento no Internet Explorer 11.
ms.date: 05/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: b8d027d4d583d42aa4efbe29e080afcd17297a74
ms.sourcegitcommit: fd04b41f513dbe9e623c212c1cbd877ae2285da0
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/11/2022
ms.locfileid: "65313210"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a>Testar seu Office suplemento no Internet Explorer 11

> [!IMPORTANT]
> **O Internet Explorer ainda é Office suplementos**
>
> Algumas combinações de plataformas e versões do Office, incluindo versões de compra única por meio do Office 2019, ainda usam o controle de modo de exibição da Web que vem com o Internet Explorer 11 para hospedar suplementos, conforme explicado em Navegadores usados por [suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md). Recomendamos (mas não exige) que você continue a dar suporte a essas combinações, pelo menos de maneira mínima, fornecendo aos usuários do seu suplemento uma mensagem de falha normal quando o suplemento é iniciado no modo de exibição da Web do Internet Explorer. Lembre-se destes pontos adicionais:
>
> - Office na Web abre mais no Internet Explorer. Consequentemente, [o AppSource](/office/dev/store/submit-to-appsource-via-partner-center) não testa mais suplementos no Office na Web usando o Internet Explorer como navegador.
> - O AppSource ainda testa combinações de versões de plataforma e área de  trabalho do Office que usam o Internet Explorer, no entanto, ele só emite um aviso quando o suplemento não dá suporte ao Internet Explorer; o suplemento não é rejeitado pelo AppSource.
> - A [Script Lab não dá](../overview/explore-with-script-lab.md) mais suporte ao Internet Explorer.

Se você planeja dar suporte a versões mais antigas do Windows e Office, seu suplemento deve funcionar no controle de navegador inserível baseado no Internet Explorer 11 (IE11). Você pode usar uma linha de comando para alternar de runtimes mais modernos usados por suplementos para o runtime do Internet Explorer 11 para esse teste. Para obter informações sobre quais versões do Windows e Office usam o controle de exibição da Web do Internet Explorer 11, consulte Navegadores usados Office [suplementos](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> O Internet Explorer 11 não oferece suporte às versões do JavaScript posteriores a ES5. Se você quiser usar a sintaxe e os recursos do ECMAScript 2015 ou posterior, terá duas opções:
>
> - Escreva seu código no ECMAScript 2015 (também chamado de ES6) ou em JavaScript posterior ou em TypeScript e, em seguida, compile seu código em JavaScript ES5 usando um compilador como [babel](https://babeljs.io/) ou [tsc](https://www.typescriptlang.org/index.html).
> - Escreva no ECMAScript 2015 ou em JavaScript posterior, mas também carregue uma biblioteca de [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) , como [core-js](https://github.com/zloirock/core-js) , que permite que o IE execute seu código.
>
> Para obter mais informações sobre essas opções, consulte [Suporte do Internet Explorer 11](../develop/support-ie-11.md).
>
> Além disso, o Internet Explorer 11 não oferece suporte a alguns recursos do HTML5, como mídia, gravação e localização. Para saber mais, confira [Determinar em runtime se o suplemento está em execução no Internet Explorer](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer).

> [!NOTE]
> - Office na Web não pode ser aberto no Internet Explorer 11, portanto, você não pode (e não precisa) testar seu suplemento no Office na Web com o Internet Explorer.
>
> - A Configuração de Segurança Aprimorada da (ESC) do Internet Explorer deve ser desativada para os suplementos Web do Office funcionarem. Se estiver usando um computador Windows Server como cliente, ao desenvolver suplementos observe se a ESC está ativada por padrão no Windows Server.

## <a name="switch-to-the-internet-explorer-11-webview"></a>Alternar para o modo de exibição da Web do Internet Explorer 11

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

Há duas maneiras de alternar o modo de exibição da Web do Internet Explorer. Você pode executar um comando simples em um prompt de comando ou instalar uma versão do Office que usa o Internet Explorer por padrão. Recomendamos o primeiro método. Mas você deve usar o segundo nos cenários a seguir.

- Seu projeto foi desenvolvido com o Visual Studio e o IIS. Não é baseado node.js segurança.
- Você quer ser absolutamente robusto em seus testes.
- Se, por algum motivo, a ferramenta de linha de comando não funcionar.

### <a name="switch-via-the-command-line"></a>Alternar por meio da linha de comando

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### <a name="install-a-version-of-office-that-uses-internet-explorer"></a>Instalar uma versão do Office que usa o Internet Explorer

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]

## <a name="see-also"></a>Confira também

* [Testar e depurar Suplementos do Office](test-debug-office-add-ins.md)
* [Realizar sideload de suplementos do Office para teste](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [Depurar os suplementos usando as ferramentas de desenvolvedor para o Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
* [Anexar um depurador do painel de tarefas](attach-debugger-from-task-pane.md)
