---
title: Evite usar o método context.sync em loops
description: Saiba como usar o loop dividido e os padrões de objetos correlacionados para evitar chamar context.sync em um loop.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6b0239e05a597949160afbb2604143f3d6626462
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958696"
---
# <a name="avoid-using-the-contextsync-method-in-loops"></a>Evite usar o método context.sync em loops

> [!NOTE]
> Este artigo pressupõe que você esteja além do estágio inicial de trabalhar com pelo menos uma das quatro APIs&mdash;JavaScript do Office específicas do aplicativo para Excel, Word, OneNote e Visio&mdash;que usam um sistema em lotes para interagir com o documento do Office. Em particular, você deve saber o que uma chamada `context.sync` faz e deve saber o que é um objeto de coleção. Se você não estiver nesse estágio, comece compreendendo a [API JavaScript do Office](../develop/understanding-the-javascript-api-for-office.md) e a documentação vinculada em "específico do aplicativo" nesse artigo.

Para alguns cenários de programação em Suplementos do Office que usam um dos modelos de API específicos do aplicativo (para Excel, Word, OneNote e Visio), seu código precisa ler, gravar ou processar alguma propriedade de cada membro de um objeto de coleção. Por exemplo, um suplemento do Excel que precisa obter os valores de cada célula em uma coluna de tabela específica ou um suplemento do Word que precisa realçar cada instância de uma cadeia de caracteres no documento. Você precisa iterar `items` sobre os membros na propriedade do objeto de coleção, mas, por motivos de desempenho, `context.sync` você precisa evitar chamar em cada iteração do loop. Cada chamada é `context.sync` uma viagem de ida e volta do suplemento para o documento do Office. Viagens de ida e volta repetidas prejudicam o desempenho, especialmente se o suplemento estiver em execução Office na Web porque as viagens de ida e volta vão pela Internet.

> [!NOTE]
> Todos os exemplos neste artigo `for` usam loops, mas as práticas descritas se aplicam a qualquer instrução de loop que possa iterar por uma matriz, incluindo o seguinte:
>
> - `for`
> - `for of`
> - `while`
> - `do while`
>
> Eles também se aplicam a qualquer método de matriz ao qual uma função é passada e aplicada aos itens na matriz, incluindo o seguinte:
>
> - `Array.every`
> - `Array.forEach`
> - `Array.filter`
> - `Array.find`
> - `Array.findIndex`
> - `Array.map`
> - `Array.reduce`
> - `Array.reduceRight`
> - `Array.some`

## <a name="writing-to-the-document"></a>Gravando no documento

No caso mais simples, você está gravando apenas para membros de um objeto de coleção, não lendo suas propriedades. Por exemplo, o código a seguir realça em amarelo todas as instâncias de "o" em um documento do Word.

> [!NOTE]
> Geralmente, é uma boa prática `context.sync` colocar um final logo antes do caractere de fechamento "}" `run` da função de aplicativo ( `Excel.run`como , `Word.run`etc.). Isso ocorre porque `run` a função faz uma chamada oculta como a `context.sync` última coisa que faz se, e somente se, houver comandos na fila que ainda não foram sincronizados. O fato de essa chamada estar oculta pode ser confuso, portanto, geralmente recomendamos que você adicione o explícito `context.sync`. No entanto, considerando que este artigo é sobre minimizar `context.sync`chamadas de , é realmente mais confuso adicionar um final totalmente desnecessário `context.sync`. Portanto, neste artigo, o deixamos de fora quando não há comandos não sincronizados no final do `run`.

```javascript
await Word.run(async function (context) {
  let startTime, endTime;
  const docBody = context.document.body;

  // search() returns an array of Ranges.
  const searchResults = docBody.search('the', { matchWholeWord: true });
  searchResults.load('font');
  await context.sync();

  // Record the system time.
  startTime = performance.now();

  for (let i = 0; i < searchResults.items.length; i++) {
    searchResults.items[i].font.highlightColor = '#FFFF00';

    await context.sync(); // SYNCHRONIZE IN EACH ITERATION
  }
  
  // await context.sync(); // SYNCHRONIZE AFTER THE LOOP

  // Record the system time again then calculate how long the operation took.
  endTime = performance.now();
  console.log("The operation took: " + (endTime - startTime) + " milliseconds.");
})
```

O código anterior levou 1 segundo completo para ser concluído em um documento com 200 instâncias de "the" no Word no Windows. Mas quando `await context.sync();` a linha dentro do loop é comentada e a mesma linha logo após o loop não ser comentado, a operação levou apenas um 1/10 de segundo. No Word na Web (com o Edge como navegador), levou 3 segundos completos com a sincronização dentro do loop e apenas 6/10 segundos com a sincronização após o loop, cerca de cinco vezes mais rápido. Em um documento com 2.000 instâncias de "the", ele levou (no Word na Web) 80 segundos com a sincronização dentro do loop e apenas 4 segundos com a sincronização após o loop, cerca de 20 vezes mais rápido.

> [!NOTE]
> Vale a pena perguntar se a versão de sincronização dentro do loop seria executada mais rapidamente se as sincronizações fosse executadas simultaneamente, `await` `context.sync()`o que poderia ser feito simplesmente removendo a palavra-chave da frente do . Isso fará com que o runtime inicie a sincronização e inicie imediatamente a próxima iteração do loop sem aguardar a conclusão da sincronização. No entanto, essa não é uma solução tão boa quanto mover `context.sync` o loop totalmente para fora por esses motivos.
>
> - Assim como os comandos em um trabalho em lotes de sincronização são enfileirados, os próprios trabalhos em lotes são enfileirados no Office, mas o Office dá suporte a não mais de 50 trabalhos em lotes na fila. Mais erros de gatilhos. Portanto, se houver mais de 50 iterações em um loop, há uma chance de que o tamanho da fila seja excedido. Quanto maior o número de iterações, maior a chance de isso acontecer.
> - "Simultaneamente" não significa simultaneamente. Ainda levaria mais tempo para executar várias operações de sincronização do que executar uma.
> - Não há garantia de que as operações simultâneas sejam concluídas na mesma ordem em que foram iniciadas. No exemplo anterior, não importa qual ordem a palavra "the" é realçada, mas há cenários em que é importante que os itens na coleção sejam processados na ordem.

## <a name="read-values-from-the-document-with-the-split-loop-pattern"></a>Ler valores do documento com o padrão de loop dividido

Evitar s `context.sync`dentro de um loop torna-se mais desafiador quando o código deve *ler uma propriedade* dos itens de coleção à medida que processa cada um. Suponha que seu código precise iterar todos os controles de conteúdo em um documento do Word e registrar o texto do primeiro parágrafo associado a cada controle. Seus instintos de programação podem levar você a executar um loop sobre os controles, `text` carregar a propriedade de cada parágrafo (primeiro), `context.sync` chamar para preencher o objeto de parágrafo proxy com o texto do documento e, em seguida, registrá-lo. Apresentamos um exemplo a seguir.

```javascript
Word.run(async (context) => {
    const contentControls = context.document.contentControls.load('items');
    await context.sync();

    for (let i = 0; i < contentControls.items.length; i++) {
      const paragraph = contentControls.items[i].getRange('Whole').paragraphs.getFirst();
      paragraph.load('text');
      await context.sync();
      console.log(paragraph.text);
    }
});
```

Nesse cenário, para evitar ter um `context.sync` loop em loop, você deve usar um padrão que chamamos de padrão **de loop dividido** . Vamos ver um exemplo concreto do padrão antes de chegarmos a uma descrição formal dele. Veja como o padrão de loop dividido pode ser aplicado ao snippet de código anterior. Observe o seguinte sobre este código.

- Agora há dois loops e o `context.sync` que vem entre eles, portanto, não há nenhum `context.sync` loop dentro de nenhum dos dois loops.
- O primeiro loop itera `text` pelos itens no objeto de coleção e carrega a propriedade da mesma forma que o loop original, mas o primeiro loop não pode registrar o texto do parágrafo porque ele não contém mais um para preencher a `text` `context.sync` `paragraph` propriedade do objeto proxy. Em vez disso, ele adiciona o `paragraph` objeto a uma matriz.
- O segundo loop itera pela matriz que foi criada pelo primeiro loop e registra o `text` de cada `paragraph` item. Isso é possível porque o `context.sync` que veio entre os dois loops preencheu todas as `text` propriedades.

```javascript
Word.run(async (context) => {
    const contentControls = context.document.contentControls.load("items");
    await context.sync();

    const firstParagraphsOfCCs = [];
    for (let i = 0; i < contentControls.items.length; i++) {
      const paragraph = contentControls.items[i].getRange('Whole').paragraphs.getFirst();
      paragraph.load('text');
      firstParagraphsOfCCs.push(paragraph);
    }

    await context.sync();

    for (let i = 0; i < firstParagraphsOfCCs.length; i++) {
      console.log(firstParagraphsOfCCs[i].text);
    }
});
```

O exemplo anterior sugere o procedimento a seguir para transformar um loop que contém um `context.sync` padrão de loop dividido.

1. Substitua o loop por dois loops.
2. Crie um primeiro loop para iterar sobre a coleção e adicionar cada item a uma matriz enquanto também carrega qualquer propriedade do item que seu código precisa ler.
3. Após o primeiro loop, chame `context.sync` para popular os objetos proxy com quaisquer propriedades carregadas.
4. Siga o `context.sync` segundo loop para iterar sobre a matriz criada no primeiro loop e ler as propriedades carregadas.

## <a name="process-objects-in-the-document-with-the-correlated-objects-pattern"></a>Processar objetos no documento com o padrão de objetos correlacionados

Vamos considerar um cenário mais complexo em que o processamento dos itens na coleção requer dados que não estão nos itens em si. O cenário prevê um suplemento do Word que opera em documentos criados com base em um modelo com algum texto clichê. Dispersas no texto estão uma ou mais instâncias das seguintes cadeias de caracteres de espaço reservado: "{Coordenador}", "{Deputy}" e "{Manager}". O suplemento substitui cada espaço reservado pelo nome de alguma pessoa. A interface do usuário do suplemento não é importante para este artigo. Por exemplo, ele pode ter um painel de tarefas com três caixas de texto, cada uma rotulada com um dos espaços reservados. O usuário insere um nome em cada caixa de texto e pressiona um **botão** Substituir. O manipulador do botão cria uma matriz que mapeia os nomes para os espaços reservados e, em seguida, substitui cada espaço reservado pelo nome atribuído.

Você não precisa realmente produzir um suplemento com essa interface do usuário para experimentar o código. Você pode usar a [ferramenta Script Lab para](../overview/explore-with-script-lab.md) criar protótipos do código importante. Use a instrução de atribuição a seguir para criar a matriz de mapeamento.

```javascript
const jobMapping = [
        { job: "{Coordinator}", person: "Sally" },
        { job: "{Deputy}", person: "Bob" },
        { job: "{Manager}", person: "Kim" }
    ];
```

O código a seguir mostra como você pode substituir cada espaço reservado pelo nome atribuído se você usou `context.sync` loops interno.

```javascript
Word.run(async (context) => {

    for (let i = 0; i < jobMapping.length; i++) {
      let options = Word.SearchOptions.newObject(context);
      options.matchWildCards = false;
      let searchResults = context.document.body.search(jobMapping[i].job, options);
      searchResults.load('items');

      await context.sync(); 

      for (let j = 0; j < searchResults.items.length; j++) {
        searchResults.items[j].insertText(jobMapping[i].person, Word.InsertLocation.replace);

        await context.sync();
      }
    }
});
```

No código anterior, há um loop externo e interno. Cada um deles contém um `context.sync`. Com base no primeiro snippet de código neste artigo, `context.sync` você provavelmente verá que o loop interno pode simplesmente ser movido após o loop interno. Mas isso ainda deixaria o código com um `context.sync` (dois deles na verdade) no loop externo. O código a seguir mostra como você pode remover `context.sync` dos loops. Discutiremos o código abaixo.

```javascript
Word.run(async (context) => {

    const allSearchResults = [];
    for (let i = 0; i < jobMapping.length; i++) {
      let options = Word.SearchOptions.newObject(context);
      options.matchWildCards = false;
      let searchResults = context.document.body.search(jobMapping[i].job, options);
      searchResults.load('items');
      let correlatedSearchResult = {
        rangesMatchingJob: searchResults,
        personAssignedToJob: jobMapping[i].person
      }
      allSearchResults.push(correlatedSearchResult);
    }

    await context.sync()

    for (let i = 0; i < allSearchResults.length; i++) {
      let correlatedObject = allSearchResults[i];

      for (let j = 0; j < correlatedObject.rangesMatchingJob.items.length; j++) {
        let targetRange = correlatedObject.rangesMatchingJob.items[j];
        let name = correlatedObject.personAssignedToJob;
        targetRange.insertText(name, Word.InsertLocation.replace);
      }
    }

    await context.sync();
});
```

Observe que o código usa o padrão de loop dividido.

- O loop externo do exemplo anterior foi dividido em dois. (O segundo loop tem um loop interno, que é esperado porque o código está iterando em um conjunto de trabalhos (ou espaços reservados) e dentro desse conjunto ele está iterando sobre os intervalos correspondentes.)
- Há um após cada `context.sync` loop principal, mas não dentro `context.sync` de nenhum loop.
- O segundo loop principal itera por meio de uma matriz criada no primeiro loop.

Mas a matriz criada no primeiro loop não contém  apenas um objeto do Office como o primeiro loop fez na seção Lendo valores do documento com o padrão [de loop dividido](#read-values-from-the-document-with-the-split-loop-pattern). Isso ocorre porque algumas das informações necessárias para processar os objetos intervalo de palavras não estão nos próprios objetos Range, mas vêm da `jobMapping` matriz.

Portanto, os objetos na matriz criada no primeiro loop são objetos personalizados que têm duas propriedades. A primeira é uma matriz de Intervalos de Palavras que correspondem a um cargo específico (ou seja, uma cadeia de caracteres de espaço reservado) e a segunda é uma cadeia de caracteres que fornece o nome da pessoa atribuída ao trabalho. Isso facilita a gravação e a leitura do loop final, pois todas as informações necessárias para processar um determinado intervalo estão contidas no mesmo objeto personalizado que contém o intervalo. O nome que deve substituir _**correlatedObject.rangesMatchingJob.items**[j]_ é a outra propriedade do mesmo objeto: _**correlatedObject.personAssignedToJob**_.

Chamamos essa variação do padrão de loop dividido do **padrão de objetos correlacionados** . A ideia geral é que o primeiro loop cria uma matriz de objetos personalizados. Cada objeto tem uma propriedade cujo valor é um dos itens em um objeto de coleção do Office (ou uma matriz desses itens). O objeto personalizado tem outras propriedades, cada uma das quais fornece informações necessárias para processar os objetos do Office no loop final. Consulte a seção [Outros exemplos desses](#other-examples-of-these-patterns) padrões para um link para um exemplo em que o objeto de correlação personalizado tem mais de duas propriedades.

Outra limitação: às vezes, é necessário mais de um loop apenas para criar a matriz de objetos correlacionados personalizados. Isso pode acontecer se você precisar ler uma propriedade de cada membro de um objeto de coleção do Office apenas para coletar informações que serão usadas para processar outro objeto de coleção. (Por exemplo, seu código precisa ler os títulos de todas as colunas em uma tabela do Excel porque o suplemento aplicará um formato numérico às células de algumas colunas com base no título dessa coluna.) Mas você sempre pode manter o `context.sync`s entre os loops, em vez de em um loop. Consulte a seção [Outros exemplos desses padrões](#other-examples-of-these-patterns) para obter um exemplo.

## <a name="other-examples-of-these-patterns"></a>Outros exemplos desses padrões

- Para obter um exemplo muito simples para o Excel `Array.forEach` que usa loops, consulte a resposta aceita para esta pergunta do Stack Overflow: É possível enfileirar mais de [um context.load antes de context.sync?](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)
- Para obter um exemplo simples do Word `Array.forEach` que usa loops`await` `async`/e não usa sintaxe, consulte a resposta aceita para esta pergunta do Stack Overflow: iterando em todos os parágrafos com controles de conteúdo com a [API JavaScript do Office](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api).
- Para obter um exemplo do Word escrito em TypeScript, consulte o Exemplo de Verificador de Estilo [angular2](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker) do Suplemento do Word, especialmente a [palavra de arquivo.document.service.ts](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts). Ele tem uma mistura de `for` e `Array.forEach` loops.
- Para obter um exemplo avançado do Word, importe [esse gist](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab) para a [Script Lab ferramentas](../overview/explore-with-script-lab.md). Para obter contexto ao usar o gist, consulte a resposta aceita para o Documento de pergunta do Stack Overflow não [sincronizado após substituir o texto](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text). Este exemplo cria um tipo de objeto de correlação personalizado que tem três propriedades. Ele usa um total de três loops para construir a matriz de objetos correlacionados e mais dois loops para fazer o processamento final. Há uma mistura de `for` e `Array.forEach` loops.
- Embora não seja estritamente um exemplo do loop dividido ou padrões de objetos correlacionados, há um exemplo avançado do Excel que mostra como converter um conjunto de valores de célula em outras moedas com apenas um único `context.sync`. Para experimentá-lo, abra [a Script Lab e](../overview/explore-with-script-lab.md) navegue até o **exemplo Conversor de Moedas Conversor**.

## <a name="when-should-you-not-use-the-patterns-in-this-article"></a>Quando você *não deve* usar os padrões neste artigo?

O Excel não pode ler mais de 5 MB de dados em uma determinada chamada de `context.sync`. Se esse limite for excedido, um erro será gerado. (Consulte a "seção suplementos do Excel" de limites de recursos e otimização de desempenho para [Suplementos do Office](resource-limits-and-performance-optimization.md#excel-add-ins) para obter mais informações.) É muito raro que esse limite seja abordado, mas se houver uma chance de que isso aconteça com o suplemento, seu código não deverá carregar todos os dados em um  único loop e seguir o loop com um `context.sync`. Mas você ainda deve evitar ter uma em `context.sync` cada iteração de um loop em um objeto de coleção. Em vez disso, defina subconjuntos dos itens na coleção e loop sobre cada subconjunto, por sua vez, com um `context.sync` entre os loops. Você pode estruturar isso com um loop externo que itera sobre os subconjuntos `context.sync` e contém o em cada uma dessas iterações externas.
