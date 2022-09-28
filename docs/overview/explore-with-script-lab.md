---
title: Explore a API JavaScript do Office usando o Script Lab
description: Use o script Lab para explorar a funcionalidade de protótipo e a API do Office JS.
ms.date: 09/27/2022
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 327810e43d87894f2850f93bd07000f1cdae98f3
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092984"
---
# <a name="explore-office-javascript-api-using-script-lab"></a>Explore a API JavaScript do Office usando o Script Lab

Os suplementos [Script Lab](https://appsource.microsoft.com/product/office/WA104380862) e [Script Lab para Outlook](https://appsource.microsoft.com/product/office/WA200001603), que estão disponíveis gratuitamente no AppSource, permitem explorar a API JavaScript do Office enquanto você trabalha em um programa do Office, como o Excel ou o Outlook. O Script Lab é uma ferramenta conveniente para adicionar ao seu kit de ferramentas de desenvolvimento durante a criação do protótipo e verificação da funcionalidade que deseja adicionar ao suplemento.

## <a name="what-is-script-lab"></a>O que é o script Lab?

O Script Lab é uma ferramenta para quem deseja saber como desenvolver suplementos do Office usando a API JavaScript do Office no Excel, no Outlook, no Word e no PowerPoint. Ele fornece IntelliSense para que você possa ver o que está disponível e que foi criado na estrutura de Mônaco, a mesma estrutura usada pelo código do Visual Studio. Por meio do Script Lab, você pode acessar uma biblioteca de amostras para experimentar rapidamente recursos ou até mesmo usar um exemplo como o ponto de partida para o seu próprio código. Você pode até usar o Script Lab para experimentar as APIs de visualização.

Parece bom? Dê uma olhada neste vídeo de um minuto para ver Script Lab em ação.

[![Visualização de vídeo mostrando o Script Lab em execução no Excel, Word e PowerPoint.](../images/screenshot-wide-youtube.png 'Visualização de vídeo do Script Lab')](https://aka.ms/scriptlabvideo)

## <a name="key-features"></a>Principais recursos

O script Lab oferece vários recursos para ajudá-lo a explorar a funcionalidade do suplemento API e protótipo do Office JavaScript.

### <a name="explore-samples"></a>Explorar amostras

Comece a trabalhar rapidamente com um conjunto de exemplos internos que mostram como concluir tarefas com a API. Você pode executar as amostras para ver instantaneamente o resultado no painel de tarefas ou documento, examinar os exemplos para saber como a API funciona e até mesmo usar amostras para criar um protótipo do seu próprio suplemento.

![Exemplos.](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a>Código e estilo

Além de código JavaScript ou TypeScript que chama a API do Office JS, cada snippet também contém marcação HTML que define o conteúdo do painel de tarefas e CSS que define a aparência do painel de tarefas. Você pode personalizar a marcação HTML e CSS para experimentar o posicionamento e o estilo de elementos à medida que você cria seu próprio suplemento no painel de tarefas.

> [!TIP]
> Para chamar APIs de visualização em um snippet, você precisará atualizar as bibliotecas do snippet para usar a CDN (rede de distribuição de conteúdo) beta (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) e as definições de tipo de visualização `@types/office-js-preview`. Além disso, algumas APIs de visualização são acessíveis apenas se você se inscreveu no programa [Office Insider](https://insider.office.com) e está executando uma compilação do Office Insider.

### <a name="save-and-share-snippets"></a>Salvar e compartilhar trechos

Por padrão, os trechos abertos no Script Lab serão salvos no cache do navegador. Para salvar um trecho permanentemente, você pode exportá-lo para um [GitHub gist](https://gist.github.com). Crie uma propriedade secreta para salvar um trecho exclusivo para seu próprio uso ou criar uma conta pública se planejar compartilhá-la com outras pessoas.

![Opções de compartilhamento.](../images/script-lab-share.jpg)

### <a name="import-snippets"></a>Importar trechos

Você pode importar um trecho para o Script Lab especificando a URL para o [do GitHub público](https://gist.github.com) onde o snippet YAML está armazenado ou colando-o no YAML completo do trecho. Esse recurso pode ser útil em situações em que outra pessoa compartilhou trechos com você publicando-o em uma oferta do GitHub ou fornecendo o YAML do trecho.

![Opção importar snippet.](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a>Clientes com suporte

O Script Lab tem suporte para o Excel, o Word e o PowerPoint nos clientes a seguir.

- Office no Windows\*
- Office 2016 ou posterior no Mac
- Office na Web

O Script Lab para Outlook está disponível para os seguintes clientes.

- Outlook no Windows\*
- Outlook 2016 ou posterior no Mac
- Outlook na Web ao usar os navegadores Chrome, Microsoft Edge ou Safari

Para obter mais detalhes sobre o Script Lab para Outlook, confira a [postagem do blog](https://devblogs.microsoft.com/microsoft365dev/script-lab-now-supports-outlook/) relacionada.

> [!IMPORTANT]
> \*Script Lab funciona mais com combinações de plataforma e versão do Office que usam o Internet Explorer para hospedar suplementos. Isso inclui versões perpétuas do Office por meio do Office 2019. Para obter mais informações, confira [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).

## <a name="next-steps"></a>Próximas etapas

Para usar o Script Lab no Excel, no Word ou no PowerPoint, instale o [suplemento do Script Lab](https://appsource.microsoft.com/product/office/WA104380862) do AppSource.

Para usar o Script Lab para Outlook, instale o suplemento [Script Lab para Outlook](https://appsource.microsoft.com/product/office/wa200001603) do AppSource.

Você é bem-vindo a expandir a biblioteca de exemplo no Script Lab, contribuindo com novos trechos para o [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) repositório do GitHub.

Quando estiver pronto para criar seu primeiro suplemento do Office, experimente o início rápido para [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](../quickstarts/outlook-quickstart.md), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md)ou [Project](../quickstarts/project-quickstart.md).

## <a name="see-also"></a>Confira também

- [Obter o Script Lab para Excel, Word ou Powerpoint](https://appsource.microsoft.com/product/office/WA104380862)
- [Script Lab para Outlook](https://appsource.microsoft.com/product/office/wa200001603)
- [Saiba mais sobre o Script Lab](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [Participe do programa para desenvolvedores do Microsoft 365](https://developer.microsoft.com/office/dev-program)
- [Desenvolvimento de Suplementos do Office ](../develop/develop-overview.md)
- [Saiba mais sobre o Programa para Desenvolvedores do Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
