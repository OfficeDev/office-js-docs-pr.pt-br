---
title: Diretrizes de ícone de estilo monolinha para Os Complementos do Office
description: Obter diretrizes para usar ícones de ícone de estilo Monoline nos Complementos do Office.
ms.date: 2/09/2021
localization_priority: Normal
ms.openlocfilehash: 262cde129c7f7d3dd3f32b32e0a8e750cf016ef8
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237949"
---
# <a name="monoline-style-icon-guidelines-for-office-add-ins"></a>Diretrizes de ícone de estilo monolinha para Os Complementos do Office

Iconografia de estilo monolinha é usada em aplicativos do Office. If you would prefer that your icons match the Fresh style of non-subscription Office 2013+, see [Fresh style icon guidelines for Office Add-ins](add-in-icons-fresh.md).

## <a name="office-monoline-visual-style"></a>Estilo visual monoline do Office

O objetivo do estilo Monoline é ter uma iconografia consistente, clara e acessível para comunicar ações e recursos com elementos visuais simples, garantir que os ícones sejam acessíveis a todos os usuários e ter um estilo consistente com os usados em outro lugar no Windows.

As diretrizes a seguir são para desenvolvedores terceirizados que querem criar ícones para recursos que serão consistentes com os ícones já presentes nos produtos do Office.

### <a name="design-principles"></a>Princípios de design

- Simples, limpo, claro.
- Conter somente os elementos necessários.
- Inspirado no estilo de ícone do Windows.
- Acessível para todos os usuários.

#### <a name="conveying-meaning"></a>Transmitir significado

- Use elementos descritivos, como uma página, para representar um documento ou um envelope para representar o email.
- Use o mesmo elemento para representar o mesmo conceito, ou seja, o email é sempre representado por um envelope, não um carimbo.
- Use uma metáfora principal durante o desenvolvimento de conceitos.

#### <a name="reduction-of-elements"></a>Redução de elementos

- Reduza o ícone ao seu significado principal, usando apenas elementos essenciais para a metáfora.
- Limite o número de elementos em um ícone a dois, independentemente do tamanho do ícone.

#### <a name="consistency"></a>Consistência

Tamanhos, organização e cor dos ícones devem ser consistentes.

#### <a name="styling"></a>Estilo

##### <a name="perspective"></a>Perspectiva

Ícones monolinha são voltados para frente por padrão. Determinados elementos que exigem perspectiva e/ou rotação, como um cubo, são permitidos, mas as exceções devem ser mantidas no mínimo.

##### <a name="embellishment"></a>Embelezamento

Monoline é um estilo mínimo limpo. Tudo usa cor plana, o que significa que não há gradientes, texturas ou fontes de luz.

## <a name="designing"></a>Projetando

### <a name="sizes"></a>Tamanhos

Recomendamos que você produza cada ícone em todos esses tamanhos para dar suporte a dispositivos de alto DPI. Os *tamanhos absolutamente* necessários são 16 px, 20 px e 32 px, pois esses são os tamanhos 100%.

**16 px, 20 px, 24 px, 32 px, 40 px, 48 px, 64 px, 80 px, 96 px**

### <a name="layout"></a>Layout

A seguir está um exemplo de layout de ícone com um modificador.

![Diagrama de ícone com modificador no canto inferior direito](../images/monolineicon1.png)  ![Diagrama do mesmo ícone com plano de fundo de grade e textos explicadores adicionados para a base, modificador, preenchimento e recorte](../images/monolineicon2.png)

#### <a name="elements"></a>Elementos

- **Base**: o conceito principal que o ícone representa. Geralmente, esse é o único elemento visual necessário para o ícone, mas, às vezes, o conceito principal pode ser aprimorado com um elemento secundário, um modificador.

- **Modificador** Qualquer elemento que sobrepõe a base; ou seja, um modificador que normalmente representa uma ação ou um status. Ele modifica o elemento base agindo como uma adição, alteração ou descritor.

![Diagrama de grade com áreas base e modificadora destacadas](../images/monolineicon3.png)

### <a name="construction"></a>Construção

#### <a name="element-placement"></a>Posicionamento do elemento

Os elementos base são colocados no centro do ícone dentro do preenchimento. Se não for possível coloca-la perfeitamente centralizada, a base deverá ser errada para a parte superior direita. No exemplo a seguir, o ícone é centralizado perfeitamente.

![Diagrama mostrando ícone perfeitamente centralizado](../images/monolineicon4.png)

No exemplo a seguir, o ícone está errando para a esquerda.

![Diagrama mostrando o ícone que erra para a esquerda por 1 px](../images/monolineicon5.png)

Os modificadores quase sempre são colocados no canto inferior direito da tela do ícone. Em alguns casos raros, os modificadores são colocados em um canto diferente. Por exemplo, se o elemento base não for reconhecível com o modificador no canto inferior direito, considere colocá-lo no canto superior esquerdo.

![Diagrama mostrando quatro ícones com o modificador no canto inferior direito e um ícone com o modificador no canto superior esquerdo](../images/monolineicon6.png)

#### <a name="padding"></a>Padding

Cada ícone de tamanho tem uma quantidade especificada de preenchimento ao redor do ícone. O elemento base permanece dentro do preenchimento, mas o modificador deve se estender até a borda da tela, estendendo-se fora do preenchimento até a borda da borda do ícone. As imagens a seguir mostram o preenchimento recomendado a ser usado para cada um dos tamanhos de ícone.

|**16px**|**20px**|**24px**|**32px**|**40px**|**48px**|**64px**|**80px**|**96px**|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|![Ícone de 16 px com preenchimento 0px](../images/monolineicon7.png)|![Ícone de 20 px com preenchimento de 1px](../images/monolineicon8.png)|![Ícone de 24 px com preenchimento de 1 px](../images/monolineicon9.png)|![Ícone de 32 px com preenchimento de 2px](../images/monolineicon10.png)|![Ícone de 40 px com preenchimento de 2px](../images/monolineicon11.png)|![Ícone de 48 px com preenchimento de 3px](../images/monolineicon12.png)|![Ícone de 64 px com preenchimento de 4 px](../images/monolineicon13.png)|![Ícone de 80 px com preenchimento de 5 px](../images/monolineicon14.png)|![Ícone de 96 px com preenchimento de 6 px](../images/monolineicon15.png)|

#### <a name="line-weights"></a>Pesos de linha

Monoline é um estilo que é substituído por formas de linha e delineadas. Dependendo do tamanho que você está produzindo, o ícone deve usar os pesos de linha a seguir.

|Tamanho do ícone:|16px|20px|24px|32px|40px|48px|64px|80px|96px|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|**Peso da linha:**|1px|1px|1px|1px|2px|2px|2px|2px|3px|
|**Ícone de exemplo:**|![Ícone de 16 px](../images/monolineicon16.png)|![Ícone de 20 px](../images/monolineicon17.png)|![Ícone de 24 px](../images/monolineicon18.png)|![Ícone de 32 px](../images/monolineicon19.png)|![Ícone de 40 px](../images/monolineicon20.png)|![Ícone de 48 px](../images/monolineicon21.png)|![Ícone de 64 px](../images/monolineicon22.png)|![Ícone de 80 px](../images/monolineicon23.png)|![Ícone de 96 px](../images/monolineicon24.png)|

#### <a name="cutouts"></a>Recortes

Quando um elemento icon é colocado sobre outro elemento, um recorte (do elemento inferior) é usado para fornecer espaço entre os dois elementos, principalmente para fins de leitura. Isso geralmente acontece quando um modificador é colocado sobre um elemento base, mas também há casos em que nenhum dos elementos é um modificador. Esses recortes entre os dois elementos às vezes são chamados de "lacuna".

O tamanho da lacuna deve ter a mesma largura que a espessura da linha usada nesse tamanho. Se você criar um ícone de 16 px, a largura do intervalo seria de 1 px e, se for um ícone de 48 px, a lacuna deve ser de 2 px. O exemplo a seguir mostra um ícone de 32 px com uma lacuna de 1 px entre o modificador e a base subjacente.

![Ícone de 32 px com uma lacuna de 1px entre o modificador e a base subjacente](../images/monolineicon25.png)

Em alguns casos, a lacuna poderá ser ampliada em 1/2 px se o modificador tiver uma borda diagonal ou curva e o intervalo padrão não fornecer separação suficiente. Isso provavelmente afetará apenas os ícones com peso de linha de 1px: 16 px, 20 px, 24 px e 32 px.

#### <a name="background-fills"></a>Preenchimentos de plano de fundo

A maioria dos ícones no conjunto de ícones monolinha exige preenchimentos de plano de fundo. No entanto, há casos em que o objeto não teria naturalmente um preenchimento, portanto, nenhum preenchimento deve ser aplicado. Os ícones a seguir têm um preenchimento branco.

![Compilação de cinco ícones com preenchimento branco](../images/monolineicon26.png)

Os ícones a seguir não têm preenchimento. (O ícone de engrenagem é incluído para mostrar que o buraco central não está preenchido.)

![Compilação de cinco ícones sem preenchimento](../images/monolineicon27.png)

##### <a name="best-practices-for-fills"></a>Práticas recomendadas para preenchimentos

###### <a name="dos"></a>Dos:

- Preencha qualquer elemento que tenha um limite definido e naturalmente tenha um preenchimento.
- Use uma forma separada para criar o preenchimento de plano de fundo.
- Use **o preenchimento de plano** de fundo da [paleta de cores.](#color)
- Mantenha a separação de pixel entre elementos sobrepostos.
- Preencher entre vários objetos.

###### <a name="donts"></a>O que não fazer:

- Não preencha objetos que não seriam preenchidos naturalmente; por exemplo, um clipe de papel.
- Não preencha colchetes.
- Não preencha por trás de números ou caracteres alfa.

### <a name="color"></a>Cor

A paleta de cores foi projetada para simplicidade e acessibilidade. Ele contém 4 cores neutras e duas variações para azul, verde, amarelo, vermelho e roxo. Laranja não está incluído intencionalmente na paleta de cores de ícone monoline. Cada cor deve ser usada de maneiras específicas, conforme descrito nesta seção.

#### <a name="palette"></a>Paleta

![Os quatro tons de cinza em monolinha: cinza escuro para contorno autônomo ou contorno, cinza médio para contorno ou conteúdo, cinza muito claro para preenchimento de plano de fundo e cinza claro para preenchimento](../images/monoline-grayshades.png)

![A paleta de cores em monoline inclui um sombreado de azul, verde, amarelo, vermelho e roxo para autônomo, contorno e preenchimento](../images/monoline-colors.png)

#### <a name="how-to-use-color"></a>Como usar cores

Na paleta de cores Monoline, todas as cores têm variações Autônomas, Contorno e Preenchimento. Geralmente, os elementos são construídos com um preenchimento e uma borda. As cores são aplicadas em um dos seguintes padrões:

- A cor autônoma apenas para objetos que não têm preenchimento.
- A borda usa a cor Outline e o preenchimento usa a cor Preenchimento.
- A borda usa a cor Autônoma e o preenchimento usa a cor Preenchimento do Plano de Fundo.

A seguir estão exemplos de uso de cor.

![Compilação de três ícones com cor em uma borda ou preenchimento ou ambos](../images/monolineicon28.png)

A situação mais comum será ter um elemento que use Cinza Escuro Autônomo com Preenchimento de Plano de Fundo.

Ao usar um preenchimento colorido, ele sempre deve estar com sua cor de contorno correspondente. Por exemplo, Preenchimento Azul só deve ser usado com Contorno Azul. Mas há duas exceções a essa regra geral:

- O preenchimento de plano de fundo pode ser usado com qualquer cor Autônoma.
- O preenchimento cinza claro pode ser usado com duas cores diferentes de contorno: cinza escuro ou cinza médio.

#### <a name="when-to-use-color"></a>Quando usar cores

A cor deve ser usada para transmitir o significado do ícone em vez de para embelezamento. Ele deve **realçar a** ação para o usuário. Quando um modificador é adicionado a um elemento base que tem cor, o elemento base é normalmente transformado em Cinza Escuro e Preenchimento de Plano de Fundo para que o modificador possa ser o elemento de cor, como o caso abaixo com o modificador "X" sendo adicionado à base de imagem no ícone mais à esquerda do conjunto a seguir.

![Compilação de cinco ícones que usam cor](../images/monolineicon29.png)

Você deve limitar seus ícones a **uma** cor adicional, que não seja a Delineada e o Preenchimento mencionados acima. No entanto, mais cores podem ser usadas se for essencial para sua metáfora, com um limite de duas cores adicionais diferentes de cinza. Em casos raros, há exceções quando mais cores são necessárias. A seguir estão bons exemplos de ícones que usam apenas uma cor.

  ![Compilação de cinco ícones que usam uma cor cada](../images/monolineicon30.png)

Mas os ícones a seguir usam muitas cores.

  ![Compilação de cinco ícones que usam várias cores cada](../images/monolineicon31.png)

Use **Cinza Médio** para o "conteúdo" interno, como linhas de grade em um ícone de uma planilha. Cores internas adicionais são usadas quando o conteúdo precisa mostrar o comportamento do controle.

![Compilação de cinco ícones com elementos interiores cinza médios](../images/monolineicon32.png)

#### <a name="text-lines"></a>Linhas de texto

Quando linhas de texto estão em um "contêiner" (por exemplo, texto em um documento), use cinza médio. As linhas de texto que não estão em um contêiner devem ser **cinza escuro.**

### <a name="text"></a>Texto

Evite usar caracteres de texto em ícones. Como os produtos do Office são usados em todo o mundo, queremos manter os ícones com o idioma neutro possível.

## <a name="production"></a>Produção

### <a name="icon-file-format"></a>Formato de arquivo de ícone

Os ícones finais devem ser salvos como arquivos de imagem .png. Use o formato PNG com um plano de fundo transparente e tenha profundidade de 32 bits.
