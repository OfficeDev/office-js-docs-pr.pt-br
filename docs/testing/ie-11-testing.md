---
title: Teste do Internet Explorer 11
description: Teste seu Office no Internet Explorer 11.
ms.date: 10/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: cfa6a35565fdca28eab9734ccde9fc8fbb2e8270
ms.sourcegitcommit: a37be80cf47a37c85b7f5cab216c160f4e905474
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/09/2021
ms.locfileid: "60250515"
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
> Além disso, o Internet Explorer 11 não oferece suporte a alguns recursos do HTML5, como mídia, gravação e localização.

> [!NOTE]
> Office na Web não pode ser aberto no Internet Explorer 11, portanto, você não pode (e não precisa) testar seu complemento no Office na Web com o Internet Explorer.

## <a name="prerequisites"></a>Pré-requisitos

- [Node.js](https://nodejs.org/) (a versão mais recente de [LTS](https://nodejs.org/about/releases))

Estas instruções pressuem que você tenha criado um projeto de gerador Yo Office antes. Se você não tiver feito isso antes, considere ler um início rápido, como este para Excel [de Excel.](../quickstarts/excel-quickstart-jquery.md)

## <a name="switching-to-the-internet-explorer-11-webview"></a>Alternando para o webview do Internet Explorer 11

1. Crie um projeto yo Office gerador. Não importa o tipo de projeto selecionado, essa ferramenta funcionará com todos os tipos de projeto.

    > [!NOTE]
    > Se você tiver um projeto existente e quiser adicionar essa ferramenta sem criar um novo projeto, pule esta etapa e vá para a próxima etapa. 

1. Na pasta raiz do seu projeto, execute o seguinte na linha de comando. Este exemplo supõe que o arquivo de manifesto do seu projeto está na raiz. Se não estiver, especifique o caminho relativo para o arquivo de manifesto. Você deve ver uma mensagem na linha de comando que o tipo de exibição da Web agora está definido como IE.

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

> [!TIP]
> Não é necessário usar esse comando, mas deve ajudar a depurar a maioria dos problemas relacionados ao tempo de execução do Internet Explorer 11. Para uma robustez completa, você deve testar o uso de computadores com várias combinações de Windows 7, 8.1, 10 e 11 e várias versões de Office. Para obter mais informações, consulte [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md) and [How to revert to an earlier version of Office](https://support.microsoft.com/topic/2bd5c457-a917-d57e-35a1-f709e3dda841).

### <a name="command-options"></a>Opções de comando

O comando também pode ter vários tempos de `office-addin-dev-settings webview` execução como argumentos:

- ie
- edge
- Padrão.

## <a name="see-also"></a>Confira também

* [Testar e depurar Suplementos do Office](test-debug-office-add-ins.md)
* [Realizar sideload de suplementos do Office para teste](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [Depurar os complementos usando ferramentas de desenvolvedor no Windows](debug-add-ins-using-f12-developer-tools-on-windows.md)
* [Anexar um depurador do painel de tarefas](attach-debugger-from-task-pane.md)
