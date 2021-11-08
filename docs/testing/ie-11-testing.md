---
title: Teste do Internet Explorer 11
description: Teste seu Office no Internet Explorer 11.
ms.date: 11/02/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8932545aa692073babeddb6ab22a213466a7c2ba
ms.sourcegitcommit: a3debae780126e03a1b566efdec4d8be83e405b8
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/03/2021
ms.locfileid: "60809032"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a>Testar seu Office de usuário no Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer ainda usado em Office de complementos**
>
> A Microsoft está encerrando o suporte para o Internet Explorer, mas isso não afeta significativamente Office Desempios. Algumas combinações de plataformas e versões Office, incluindo versões de compra única por meio do Office 2019, continuarão a usar o controle webview que vem com o Internet Explorer 11 para hospedar os complementos, conforme explicado em Navegadores usados por Office [Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Além disso, o suporte a essas combinações e, portanto, para o Internet Explorer, ainda é necessário para os complementos enviados ao [AppSource](/office/dev/store/submit-to-appsource-via-partner-center). Duas coisas *estão mudando:*
>
> - Office na Web abre mais no Internet Explorer. Consequentemente, o AppSource não testa mais os Office na Web usando o Internet Explorer como navegador. Mas o AppSource ainda testa combinações de plataforma e Office *desktop* que usam o Internet Explorer.
> - A [Script Lab não](../overview/explore-with-script-lab.md) dá mais suporte ao Internet Explorer.

Se você planeja comercializar seu complemento por meio do AppSource ou planeja dar suporte a versões mais antigas do Windows e Office, o seu complemento deve funcionar no controle de navegador in-loca que se baseia no Internet Explorer 11 (IE11). Você pode usar uma linha de comando para alternar de tempos de execução mais modernos usados pelos complementos para o tempo de execução do Internet Explorer 11 para esse teste. Para obter informações sobre quais versões do Windows e Office usam o controle de exibição da Web do Internet Explorer 11, consulte Navegadores usados por [Office Dep.](../concepts/browsers-used-by-office-web-add-ins.md)

> [!IMPORTANT]
> O Internet Explorer 11 não oferece suporte às versões do JavaScript posteriores a ES5. Se você quiser usar a sintaxe e os recursos do ECMAScript 2015 ou posterior, você tem duas opções:
>
> - Escreva seu código no ECMAScript 2015 (também chamado de ES6) ou javaScript posterior ou em TypeScript e compile seu código para JavaScript do ES5 usando um compilador como [o babel](https://babeljs.io/) ou [o tsc](https://www.typescriptlang.org/index.html).
> - Escreva em ECMAScript 2015 ou posterior JavaScript, mas também carregue uma biblioteca de [polifilamento,](https://en.wikipedia.org/wiki/Polyfill_(programming)) como [core-js,](https://github.com/zloirock/core-js) que permite ao IE executar seu código.
>
> Para obter mais informações sobre essas opções, consulte [Support Internet Explorer 11](../develop/support-ie-11.md).
>
> Além disso, o Internet Explorer 11 não oferece suporte a alguns recursos do HTML5, como mídia, gravação e localização. Para saber mais, consulte [Determine at runtime if the add-in is running in Internet Explorer](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer).

> [!NOTE]
> Office na Web não pode ser aberto no Internet Explorer 11, portanto, você não pode (e não precisa) testar seu complemento no Office na Web com o Internet Explorer.

## <a name="switch-to-the-internet-explorer-11-webview"></a>Alternar para o Webview do Internet Explorer 11

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

Há duas maneiras de alternar o Webview do Internet Explorer. Você pode executar um comando simples em um prompt de comando ou instalar uma versão do Office que usa o Internet Explorer por padrão. Recomendamos o primeiro método. Mas você deve usar o segundo nos cenários a seguir.

- Seu projeto foi desenvolvido com Visual Studio e IIS. Não é baseado em node.js.
- Você deseja ser absolutamente robusto em seus testes.
- Se, por qualquer motivo, a ferramenta de linha de comando não funcionar.

### <a name="switch-via-the-command-line"></a>Alternar pela linha de comando

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### <a name="install-a-version-of-office-that-uses-internet-explorer"></a>Instalar uma versão do Office que usa o Internet Explorer

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]

## <a name="see-also"></a>Confira também

* [Testar e depurar Suplementos do Office](test-debug-office-add-ins.md)
* [Realizar sideload de suplementos do Office para teste](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [Depurar os suplementos usando as ferramentas de desenvolvedor para o Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
* [Anexar um depurador do painel de tarefas](attach-debugger-from-task-pane.md)
