---
title: Diretrizes de tipografia para suplementos do Office
description: Saiba quais tipos e tamanhos de fonte usar em Office Add-ins.
ms.date: 05/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: f63d4b6816b916dc52711a8f4b11e826efd58105
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63742959"
---
# <a name="typography"></a>Tipografia

Segoe é o tipo de fonte padrão para o Office. Use-a no suplemento para alinhar objetos de conteúdo, caixas de diálogo e painéis de tarefas do Office. [O Fabric Core](fabric-core.md) oferece acesso a Segoe. Ele fornece um conjunto completo da fonte Segoe com muitas variações (incluindo espessura e tamanho da fonte) em classes CSS convenientes. Nem todos os tamanhos e pesos do Fabric Core serão ótimos em um Office Add-in. Para se ajustar harmoniosamente ou evitar conflitos, considere usar um subconjunto da rampa de tipo Fabric Core. A tabela a seguir lista as classes base do Fabric Core que recomendamos para uso em Office de complementos.

> [!NOTE]
> A cor do texto não está incluída nessas classes base. Use a "primária neutra" do Fabric Core para a maioria dos textos em plano de fundo branco.
>
> Para saber mais sobre tipografia disponível, consulte [Tipografia da Web](https://developer.microsoft.com/fluentui#/styles/web/typography).

|Tipo |Classe |Tamanho |Peso |Uso recomendado |
|------ |----- |---- |------ |----------------- |
|Destaque|.ms-font-xxl |28 px | Segoe Light |<ul><li>Essa classe é maior do que todos os outros elementos tipográficos no Office. Use-a com moderação para não prejudicar o ajuste na hierarquia visual.</li><li>Evite o uso de cadeias de caracteres longas em espaços restritos.</li><li>Deixe bastante espaço em branco ao redor do texto ao usar esta classe.</li><li>Comumente usada para mensagens da tela de apresentação, elementos Hero ou outras chamadas à ação.</li></ul> |
|Título|.ms-font-xl |21 px |Segoe Light | <ul><li>Essa classe corresponde ao título do painel de tarefas dos aplicativos do Office.</li><li>Use-a com moderação para evitar uma hierarquia tipográfica monótona.</li><li>Comumente usado como o elemento de nível superior, como títulos de conteúdo, página ou caixa de diálogo.</li></ul> |
|Subtítulo|.ms-font-l |17 px |Segoe Semilight | <ul><li>Essa classe é a primeira abaixo de títulos.</li><li>Comumente usada como um subtítulo, um elemento de navegação ou um cabeçalho de grupo.</li><ul> |
|Corpo|.ms-font-m |14 px |Segoe Regular |<ul><li>Muito usada como corpo de texto dentro de suplementos.</li><ul>|
|Legenda|.ms-font-xs |11 px | Segoe Regular |<ul><li>Muito usada em texto secundário ou terciário, como carimbos de data/hora, linhas, títulos ou rótulos de campo.</li><ul>|
|Annotation|.ms-font-mi |10 px |Segoe Semibold |<ul><li>A menor etapa no painel de tipos deve ser usada raramente. Está disponível para situações em que a legibilidade não é necessária.</li><ul>|
