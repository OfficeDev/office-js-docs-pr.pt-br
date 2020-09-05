---
title: Diretrizes de tipografia para suplementos do Office
description: Saiba o que são os tipos de fonte e de fontes a serem usados nos suplementos do Office.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: d7347e2e6ee01386d631fea8c2b388ad5b61005e
ms.sourcegitcommit: 10463841a977e9b8415362a3ae91b0ae5eebbf89
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/04/2020
ms.locfileid: "47399561"
---
# <a name="typography"></a>Tipografia

Segoe é o tipo de fonte padrão para o Office. Use-a no suplemento para alinhar objetos de conteúdo, caixas de diálogo e painéis de tarefas do Office. O Office UI Fabric lhe dá acesso à fonte Segoe. Ele fornece um conjunto completo da fonte Segoe com muitas variações (incluindo espessura e tamanho da fonte) em classes CSS convenientes. Nem todos os tamanhos e espessuras do Office UI Fabric terão boa aparência em um suplemento do Office. Para obter um ajuste harmonioso ou evitar conflitos, considere o uso de um subconjunto do conjunto de fontes do Fabric. A tabela a seguir lista as classes base da malha recomendadas para uso em suplementos do Office.

> [!NOTE]
> A cor do texto não está incluída nessas classes base. Use a opção "Neutro principal" do Fabric para a maioria dos textos em fundos brancos.
>
> Para saber mais sobre a tipografia disponível, consulte [Web Typography](https://developer.microsoft.com/fluentui#/styles/web/typography).

|Tipo |Classe |Tamanho |Peso |Uso recomendado |
|------ |----- |---- |------ |----------------- |
|Destaque|.ms-font-xxl |28 px | Segoe Light |<ul><li>Essa classe é maior do que todos os outros elementos tipográficos no Office. Use-a com moderação para não prejudicar o ajuste na hierarquia visual.</li><li>Evite o uso de cadeias de caracteres longas em espaços restritos.</li><li>Deixe bastante espaço em branco ao redor do texto ao usar esta classe.</li><li>Comumente usada para mensagens da tela de apresentação, elementos Hero ou outras chamadas à ação.</li></ul> |
|Título|.ms-font-xl |21 px |Segoe Light | <ul><li>Essa classe corresponde ao título do painel de tarefas dos aplicativos do Office.</li><li>Use-a com moderação para evitar uma hierarquia tipográfica monótona.</li><li>Comumente usado como o elemento de nível superior, como títulos de conteúdo, página ou caixa de diálogo.</li></ul> |
|Subtítulo|.ms-font-l |17 px |Segoe Semilight | <ul><li>Essa classe é a primeira abaixo de títulos.</li><li>Comumente usada como um subtítulo, um elemento de navegação ou um cabeçalho de grupo.</li><ul> |
|Body|.ms-font-m |14 px |Segoe Regular |<ul><li>Muito usada como corpo de texto dentro de suplementos.</li><ul>|
|Legenda|.ms-font-xs |11 px | Segoe Regular |<ul><li>Muito usada em texto secundário ou terciário, como carimbos de data/hora, linhas, títulos ou rótulos de campo.</li><ul>|
|Annotation|.ms-font-mi |10 px |Segoe Semibold |<ul><li>A menor etapa no painel de tipos deve ser usada raramente. Está disponível para situações em que a legibilidade não é necessária.</li><ul>|
