---
title: Explorar a API JavaScript do Office usando o script Lab
description: Use o script Lab para explorar a API do Office JS e a funcionalidade de protótipo.
ms.topic: article
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 9ef81443fade450a7bea519a99cb607c320dd4f6
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128640"
---
# <a name="explore-office-javascript-api-using-script-lab"></a>Explorar a API JavaScript do Office usando o script Lab

O [suplemento de laboratório de script](https://store.office.com/app.aspx?assetid=WA104380862), que está disponível gratuitamente na Office Store, permite explorar a API JavaScript do Office enquanto você estiver trabalhando em um programa do Office, como Excel ou Word. O script Lab é uma ferramenta conveniente para adicionar ao seu kit de ferramentas de desenvolvimento conforme você protótipo e verificar a funcionalidade desejada no seu suplemento.

## <a name="what-is-script-lab"></a>O que é o script Lab?

O script Lab é uma ferramenta para qualquer pessoa que deseje saber como desenvolver suplementos do Office usando a API JavaScript do Office no Excel, no Word ou no PowerPoint. Ele fornece o IntelliSense para que você possa ver o que está disponível e foi criado na estrutura de Mônaco, a mesma estrutura usada pelo Visual Studio Code. Por meio do laboratório de scripts, você pode acessar uma biblioteca de exemplos para experimentar rapidamente recursos ou pode usar um exemplo como ponto de partida para seu próprio código. Você pode até mesmo usar o script Lab para experimentar as APIs de visualização.

Parece bom até agora? Dê uma olhada neste vídeo de um minuto para ver o script Lab em ação.

[![Visualizar vídeo mostrando o laboratório de script em execução no Excel, Word e PowerPoint.] (../images/screenshot-wide-youtube.png 'Vídeo do script Lab Preview')](https://aka.ms/scriptlabvideo)

## <a name="key-features"></a>Principais recursos

O script Lab oferece vários recursos para ajudá-lo a explorar a API JavaScript do Office e a funcionalidade do suplemento de protótipo.

### <a name="explore-samples"></a>Explorar exemplos

Comece rapidamente com uma coleção de trechos de código internos que mostram como concluir determinadas tarefas com a API. Você pode executar os exemplos para ver instantaneamente o resultado no painel de tarefas ou no documento, examinar os exemplos para saber como a API funciona e até mesmo usar trechos de exemplo como base para a funcionalidade de protótipo do seu próprio suplemento.

![Exemplos](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a>Código e estilo

Além do código JavaScript ou TypeScript que chama a API do Office JS, cada trecho também contém marcação HTML que define o conteúdo do painel de tarefas e o CSS que define a aparência do painel de tarefas. Você pode personalizar a marcação HTML e o CSS para testar o posicionamento e o estilo do elemento conforme o design do painel de tarefas do protótipo para seu próprio suplemento.

> [!TIP]
> Para chamar APIs de visualização dentro de um trecho de código, você precisará atualizar as bibliotecas do trecho de código para`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`usar a CDN beta () `@types/office-js-preview`e as definições de tipo de visualização. Além disso, algumas APIs de visualização são acessíveis somente se você se inscreveu no [programa Office](https://products.office.com/office-insider) Insider e está executando uma compilação do Office Insider.

### <a name="save-and-share-snippets"></a>Salvar e compartilhar trechos de código

Por padrão, os trechos de código abertos no laboratório de script serão salvos no cache do navegador. Para salvar um trecho permanentemente, você pode exportá-lo para um [GitHub](https://gist.github.com). Crie uma propriedade secreta para salvar um trecho de código exclusivamente para uso próprio ou criar uma propriedade compartilhada se você planeja compartilhá-la com outras pessoas.

![Opções de compartilhamento](../images/script-lab-share.jpg)

### <a name="import-snippets"></a>Importar trechos

Você pode importar um trecho para o laboratório de script especificando a URL para o membro do [GitHub](https://gist.github.com) público onde o YAML de trecho de código está armazenado ou colando no YAML completo para o trecho de código. Esse recurso pode ser útil em situações em que alguém compartilhou seus trechos de código com você publicando-o em um próprio GitHub ou fornecendo a YAML de seus trechos de código.

![Opção importar trecho](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a>Clientes com suporte

O script Lab é compatível com Excel, Word e PowerPoint nos seguintes clientes.

- Office 2013 ou posterior no Windows
- Office 2016 ou posterior no Mac
- Office na Web

## <a name="next-steps"></a>Próximas etapas

Você é bem-vindo à expansão da biblioteca de exemplo no laboratório de scripts, contribuindo novos trechos de código para o repositório do GitHub [Office-js-Snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) .

Quando estiver pronto para criar seu suplemento do Office, confira o [início rápido de 5 minutos](/office/dev/add-ins/#5-minute-quick-starts) para seu aplicativo preferido do Office.

## <a name="see-also"></a>Confira também

- [Obter o laboratório de scripts](https://store.office.com/app.aspx?assetid=WA104380862)
- [Saiba mais sobre o script Lab](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [Inscreva-se no programa dev](https://developer.microsoft.com/office/dev-program)
