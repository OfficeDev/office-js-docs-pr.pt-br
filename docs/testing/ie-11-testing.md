---
title: Teste do Internet Explorer 11
description: Teste seu Office no Internet Explorer 11.
ms.date: 06/18/2021
localization_priority: Normal
ms.openlocfilehash: 8579a37f1ea48d511010b8c55cfe9fad5aa6b41acee85b1da426e25083287655
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57090123"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a>Testar seu Office de usuário no Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer ainda usado em Office de complementos**
>
> A Microsoft está encerrando o suporte para o Internet Explorer, mas isso não afeta significativamente Office Desempios. Algumas combinações de plataformas e versões Office, incluindo todas as versões de compra única por meio do Office 2019, continuarão a usar o controle webview que vem com o Internet Explorer 11 para hospedar os complementos, conforme explicado em Navegadores usados por Office [Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Além disso, o suporte a essas combinações e, portanto, para o Internet Explorer, ainda é necessário para os complementos enviados ao [AppSource](/office/dev/store/submit-to-appsource-via-partner-center). Duas coisas *estão mudando:*
>
> - O AppSource não testa mais os Office na Web usando o Internet Explorer como navegador. Mas o AppSource ainda testa combinações de plataforma e Office *desktop* que usam o Internet Explorer.
> - A [Script Lab de usuário](../overview/explore-with-script-lab.md) para de funcionar no Internet Explorer em algum momento de 2021.

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
> Para testar seu complemento no navegador do Internet Explorer 11, abra o Office na Web no Internet Explorer e coloque o [sideload do add-in](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).

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
> Não é necessário usar esse comando, mas deve ajudar a depurar a maioria dos problemas relacionados ao tempo de execução do Internet Explorer 11. Para uma robustez completa, você deve testar o uso de computadores com várias combinações de Windows 7, 8.1 e 10 e várias versões de Office. Para obter mais informações, consulte [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md) and [How to revert to an earlier version of Office](https://support.microsoft.com/topic/how-to-revert-to-an-earlier-version-of-office-2bd5c457-a917-d57e-35a1-f709e3dda841).

### <a name="command-options"></a>Opções de comando

O comando também pode ter vários tempos de `office-addin-dev-settings webview` execução como argumentos:

- ie
- edge
- Padrão.

## <a name="see-also"></a>Confira também

* [Testar e depurar Suplementos do Office](test-debug-office-add-ins.md)
* [Realizar sideload de suplementos do Office para teste](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [Depurar suplementos usando as ferramentas de desenvolvedor no Windows 10](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [Anexar um depurador do painel de tarefas](attach-debugger-from-task-pane.md)
