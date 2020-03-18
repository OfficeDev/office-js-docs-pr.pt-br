---
title: Diretrizes de ícone de estilo monoline para suplementos do Office
description: Obter diretrizes para usar ícones de ícone de estilo monoline em suplementos do Office.
ms.date: 12/09/2019
localization_priority: Normal
ms.openlocfilehash: 8d1ac2dff76b852cd69b38bd2c138d1ba43f203c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718598"
---
# <a name="monoline-style-icon-guidelines-for-office-add-ins"></a>Diretrizes de ícone de estilo monoline para suplementos do Office

O estilo monoline iconografia são usados no Office 365. Se você preferir que seus ícones correspondam ao novo estilo de não assinatura do Office 2013 +, confira [diretrizes de ícone de estilo atualizado para suplementos do Office](add-in-icons-fresh.md).

## <a name="office-monoline-visual-style"></a>Estilo visual monoline do Office

O objetivo do estilo de monolinha ter um iconografia consistente, claro e acessível para comunicar ações e recursos com visuais simples, garantir que os ícones estejam acessíveis a todos os usuários e ter um estilo consistente com aqueles usados em qualquer lugar no Windows.

As diretrizes a seguir são para desenvolvedores de terceiros que desejam criar ícones para recursos que serão consistentes com os ícones já presentes nos produtos do Office.

### <a name="design-principles"></a>Princípios de design

-   Simples, limpo, claro.
-   Conter apenas elementos necessários.
-   Estilo de ícone do Windows inspirado.
-   Acessível a todos os usuários.

#### <a name="conveying-meaning"></a>Transmitir significado

-   Use elementos descritivos, como uma página para representar um documento ou envelope para representar emails.
-   Use o mesmo elemento para representar o mesmo conceito, ou seja, mail é sempre representado por um envelope, não um carimbo.
-   Use uma metáfora principal durante o desenvolvimento do conceito.

#### <a name="reduction-of-elements"></a>Redução dos elementos

-   Reduza o ícone ao seu significado principal, usando apenas os elementos essenciais para a metáfora.
-   Limitar o número de elementos em um ícone a dois, independentemente do tamanho do ícone.

#### <a name="consistency"></a>Consistência

Os tamanhos, a organização e a cor dos ícones devem ser consistentes.

#### <a name="styling"></a>Estilo

##### <a name="perspective"></a>Perspectiva

Os ícones monoline estão voltados para o avanço por padrão. Determinados elementos que exigem perspectiva e/ou rotação, como um cubo, são permitidos, mas as exceções devem ser mantidas no mínimo.

##### <a name="embellishment"></a>Ornamento

Monolinha é um estilo mínimo limpo. Tudo usa cor plana, o que significa que não há gradientes, texturas ou fontes de luz.

## <a name="designing"></a>Planejamento

### <a name="sizes"></a>Coincidi

Recomendamos que você produza cada ícone em todos esses tamanhos para suportar dispositivos DPI alto. Os tamanhos absolutamente *exigidos* são 16px, 20px e medianiz 32px, já que são os tamanhos 100%.

**16px, 20px, medianiz 24px, medianiz 32px, 40px, 48px, 64px, 80px, 96px**

### <a name="layout"></a>Layout

Veja a seguir um exemplo de layout de ícone com um modificador.

![Exemplo de ícone com modificador](../images/monolineicon1.png)  ![O mesmo exemplo com textos explicativos de plano de fundo de grade para base, modificador, enchimento e recorte.](../images/monolineicon2.png)

#### <a name="elements"></a>Elementos

- **Base**: o conceito principal que o ícone representa. Isso geralmente é o único Visual necessário para o ícone, mas às vezes o conceito principal pode ser aprimorado com um elemento secundário, um modificador.

- **Modificador** Qualquer elemento que sobrepõe a base; ou seja, um modificador que normalmente representa uma ação ou um status. Ele modifica o elemento base agindo como uma adição, alteração ou descritor.

![Grade com as áreas de área base e modificador.](../images/monolineicon3.png)

### <a name="construction"></a>Construção

#### <a name="element-placement"></a>Posicionamento do elemento

Os elementos base são colocados no centro do ícone dentro do preenchimento. Se ele não puder ser colocado perfeitamente centralizado, a base deverá ter um erro no canto superior direito. No exemplo a seguir, o ícone está perfeitamente centralizado:

![Imagem mostrando o ícone perfeitamente centralizado](../images/monolineicon4.png)

No exemplo a seguir, o ícone é erring à esquerda.

![Imagem mostrando o ícone que ERRs à esquerda](../images/monolineicon5.png)

Modificadores quase sempre são colocados no canto inferior direito da tela de ícones. Em alguns casos raros, os modificadores são colocados em um canto diferente. Por exemplo, se o elemento base não puder ser reconhecível com o modificador no canto inferior direito, considere colocá-lo no canto superior esquerdo.

![Imagem mostrando alguns ícones com o modificador no canto inferior direito, mas um com o modificador no canto superior esquerdo](../images/monolineicon6.png)

#### <a name="padding"></a>Padding

Cada ícone de tamanho tem uma quantidade especificada de preenchimento em torno do ícone. O elemento base permanece dentro do preenchimento, mas o modificador deve arredondar para a borda da tela, estendendo-o para fora do preenchimento---para a borda da borda do ícone. As imagens a seguir mostram o preenchimento recomendado a ser usado para cada um dos tamanhos de ícone.

|**16px**|**20px**|**24px**|**32px**|**40px**|**48px**|**64px**|**80px**|**96px**|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|![ícone 16 px](../images/monolineicon7.png)|![ícone de 20 px](../images/monolineicon8.png)|![ícone de 24 px](../images/monolineicon9.png)|![ícone da 32 px](../images/monolineicon10.png)|![ícone da 40 px](../images/monolineicon11.png)|![ícone da 48 px](../images/monolineicon12.png)|![ícone da 64 px](../images/monolineicon13.png)|![ícone da 80 px](../images/monolineicon14.png)|![ícone da 96 px](../images/monolineicon15.png)|

#### <a name="line-weights"></a>Espessuras de linha

Monolinha é um estilo dominado por formas de linha e contorno. Dependendo de qual tamanho você está produzindo, o ícone deve usar os pesos de linha a seguir.

|**Tamanho do ícone:**|**16px**|**20px**|**24px**|**32px**|**40px**|**48px**|**64px**|**80px**|**96px**|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|**Espessura da linha:**|1px|1px|1px|1px|2px|2px|2px|2px|3px|
||![ícone 16 px](../images/monolineicon16.png)|![ícone de 20 px](../images/monolineicon17.png)|![ícone de 24 px](../images/monolineicon18.png)|![ícone da 32 px](../images/monolineicon19.png)|![ícone da 40 px](../images/monolineicon20.png)|![ícone da 48 px](../images/monolineicon21.png)|![ícone da 64 px](../images/monolineicon22.png)|![ícone da 80 px](../images/monolineicon23.png)|![ícone da 96 px](../images/monolineicon24.png)|

#### <a name="cutouts"></a>Recortes

Quando um elemento Icon é colocado na parte superior de outro elemento, um recorte (do elemento inferior) é usado para fornecer espaço entre os dois elementos, principalmente para fins de legibilidade. Isso geralmente ocorre quando um modificador é colocado na parte superior de um elemento base, mas também há casos em que nenhum dos elementos é um modificador. Esses recortes entre os dois elementos são, às vezes, chamados de "Gap".

O tamanho da lacuna deve ter a mesma largura que a espessura da linha usada nesse tamanho. Se estiver fazendo um ícone de 16px, a largura do espaço seria 1 px e, se for um ícone 48px, a lacuna deverá ser 2 px. O exemplo a seguir mostra um ícone medianiz 32px com uma lacuna de 1 px entre o modificador e a base subjacente.

![medianiz 32px com uma lacuna de 1 px entre o modificador e a base de base](../images/monolineicon25.png)

Em alguns casos, a lacuna pode ser aumentada em 1/2 px se o modificador tiver uma borda diagonal ou curva e a lacuna padrão não fornecer separação suficiente. Isso provavelmente afetará somente os ícones com espessura de linha 1 px; 16px, 20px, medianiz 24px e medianiz 32px.

#### <a name="background-fills"></a>Preenchimentos de plano de fundo

A maioria dos ícones no conjunto de ícones monoline exige preenchimentos de plano de fundo. No entanto, há casos em que o objeto não teria um preenchimento naturalmente, portanto, nenhum preenchimento deve ser aplicado. Os ícones a seguir têm um preenchimento branco:

![Cinco ícones têm um preenchimento branco](../images/monolineicon26.png)

Os ícones a seguir não têm preenchimento. (O ícone de engrenagem é incluído para mostrar que o orifício central não está preenchido.) ![Cinco ícones sem preenchimento](../images/monolineicon27.png)

##### <a name="best-practices-for-fills"></a>Práticas recomendadas para preenchimentos

###### <a name="dos"></a>Ataque

- Preencha qualquer elemento que tenha um limite definido e, naturalmente, teria um preenchimento.
- Use uma forma separada para criar o preenchimento do plano de fundo.
- Usar **preenchimento de plano de fundo** da [paleta de cores](#color).
- Manter a separação de pixels entre elementos sobrepostos.
- Preencher entre vários objetos.

###### <a name="donts"></a>Permitido

- Não preencha objetos que não seriam naturalmente preenchidos; por exemplo, um clipe de clipe.
- Não preencha os colchetes.
- Não preencha números ou caracteres alfabéticos.

### <a name="color"></a>Cor

A paleta de cores foi projetada para simplificar e acessibilidade. Ele contém 4 cores neutras e duas variações de azul, verde, amarelo, vermelho e roxo. A cor laranja não é incluída intencionalmente na paleta de cores do ícone monoline. Cada cor deve ser usada de formas específicas, conforme descrito nesta seção.

#### <a name="palette"></a>Paleta

![Quatro tonalidades de cinza em monolinha](../images/monoline-grayshades.png)

![A paleta de cores em monoline](../images/monoline-colors.png)

#### <a name="how-to-use-color"></a>Como usar cores

Na paleta de cores monoline, todas as cores têm variações autônomas, de estrutura de tópicos e de preenchimento. Geralmente, os elementos são construídos com um preenchimento e uma borda. As cores são aplicadas em um dos seguintes padrões:

- A cor autônoma sozinho para objetos que não têm preenchimento.
- A borda usa a cor de contorno e o preenchimento usa a cor de preenchimento.
- A borda usa a cor autônoma e o preenchimento usa a cor de preenchimento de plano de fundo.

A seguir estão exemplos de como usar cores.

![Três ícones com cor em uma borda ou preenchimento ou ambos](../images/monolineicon28.png)

A situação mais comum será ter um elemento usando cinza escuro autônomo com preenchimento de plano de fundo.

Ao usar um preenchimento colorido, ele sempre deve estar com sua cor de contorno correspondente. Por exemplo, o preenchimento azul deve ser usado apenas com o contorno azul. Mas há duas exceções a essa regra geral:

- O preenchimento de plano de fundo pode ser usado com qualquer cor independente.
- O preenchimento cinza claro pode ser usado com duas cores de contorno diferentes: cinza escuro ou cinza médio.

#### <a name="when-to-use-color"></a>Quando usar cores

A cor deve ser usada para transmitir o significado do ícone, em vez de um ornamento. Ela deve **realçar a ação** para o usuário. Quando um modificador é adicionado a um elemento base que tem cor, o elemento base é normalmente transformado em cinza escuro e preenchimento de plano de fundo para que o modificador possa ser o elemento de cor, como o caso abaixo com o modificador "X" sendo adicionado à base da imagem na extrema esquerda ícone do conjunto a seguir.

![Cinco ícones que usam cores](../images/monolineicon29.png)

Você deve limitar seus ícones a **uma** cor adicional, diferente da estrutura de tópicos e do preenchimento mencionados acima. No entanto, é possível usar mais cores se for vital para a metáfora, com um limite de duas cores adicionais além de cinza. Em casos raros, há exceções quando são necessárias mais cores. Estes são bons exemplos de ícones que usam apenas uma cor.

  ![Uma imagem de cinco ícones com uma cor cada](../images/monolineicon30.png)

Mas os ícones a seguir usam muitas cores.

  ![Uma imagem de cinco ícones com várias cores](../images/monolineicon31.png)


Use **cinza médio** para "conteúdo" interno, como linhas de grade em um ícone de uma planilha. Cores interiores adicionais são usadas quando o conteúdo precisa mostrar o comportamento do controle.

![Cinco ícones com elementos interiores de cinza médio](../images/monolineicon32.png)

#### <a name="text-lines"></a>Linhas de texto

Quando as linhas de texto estão em um "contêiner" (por exemplo, texto em um documento), use cinza médio. As linhas de texto que não estão em um contêiner devem ser **cinza escuro**.

### <a name="text"></a>Texto

Evite usar caracteres de texto em ícones. Como os produtos do Office são usados em todo o mundo, desejamos manter os ícones da forma mais neutra possível.

## <a name="production"></a>Produção

### <a name="icon-file-format"></a>Formato de arquivo de ícone

Os ícones finais devem ser salvos como arquivos de imagem. png. Use o formato PNG com um plano de fundo transparente e tenha profundidade de 32 bits.
