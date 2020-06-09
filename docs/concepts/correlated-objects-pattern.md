---
title: Evite usar o método context.sync em loops
description: Saiba como usar o loop de divisão e os padrões de objetos correlacionados para evitar chamar Context. Sync em um loop.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: bdb7340b999d74baf200aafda2d0f2f41420bd14
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608031"
---
# <a name="avoid-using-the-contextsync-method-in-loops"></a>Evite usar o método context.sync em loops

> [!NOTE]
> Este artigo pressupõe que você está além do estágio inicial do trabalho com pelo menos uma das quatro APIs JavaScript do Office específicas do host &mdash; para Excel, Word, OneNote e Visio &mdash; que usam um sistema de lote para interagir com o documento do Office. Em particular, você deve saber o que é uma chamada `context.sync` e deve saber o que é um objeto de coleção. Se você não estiver nesse estágio, comece a [entender a API JavaScript do Office](../develop/understanding-the-javascript-api-for-office.md) e a documentação vinculada a em "específico do host", neste artigo.

Para alguns cenários de programação nos suplementos do Office que usam um dos modelos de API específicos do host (para Excel, Word, OneNote e Visio), seu código precisa ler, gravar ou processar algumas propriedades de cada membro de um objeto coleção. Por exemplo, um suplemento do Excel que precisa obter os valores de cada célula de uma determinada coluna de tabela ou um suplemento do Word que precisa realçar cada instância de uma cadeia de caracteres no documento. Você precisa iterar sobre os membros na `items` Propriedade do objeto coleção; mas, por motivos de desempenho, você precisa evitar chamadas `context.sync` em cada iteração do loop. Cada chamada de `context.sync` é uma viagem de ida e volta do suplemento para o documento do Office. Repetidas viagens de ida e volta o desempenho, especialmente se o suplemento estiver sendo executado no Office na Web, pois os ciclos de ida e volta passam pela Internet.

> [!NOTE]
> Todos os exemplos neste artigo usam `for` loops, mas as práticas descritas aplicam-se a qualquer instrução de loop que possa percorrer uma matriz, incluindo o seguinte:
>
> - `for`
> - `for of`
> - `while`
> - `do while`
> 
> Eles também se aplicam a qualquer método de matriz para o qual uma função é passada e aplicada aos itens na matriz, incluindo o seguinte:
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

## <a name="writing-to-the-document"></a>Gravação no documento

No caso mais simples, você está apenas gravando para membros de um objeto de coleção, não lendo suas propriedades. Por exemplo, o código a seguir é realçado em amarelo a cada instância de "a" em um documento do Word. 

> [!NOTE]
> Geralmente, é uma boa prática ter um ponto final `context.sync` antes do caractere de fechamento "}" do `run` método host (como `Excel.run` , `Word.run` etc.). Isso ocorre porque o `run` método faz uma chamada oculta do `context.sync` como a última coisa que faz se, e somente se, houver comandos em fila que ainda não tenham sido sincronizados. O fato de esta chamada ser oculta pode ser confuso, portanto, geralmente recomendamos que você adicione o explícito `context.sync` . No entanto, Considerando que este artigo está prestes a minimizar as chamadas `context.sync` , é, na verdade, mais confuso de adicionar um final totalmente desnecessário `context.sync` . Portanto, neste artigo, deixamos de existir quando não há comandos não sincronizados no final do `run` . 

```javascript
Word.run(async function (context) {
    let startTime, endTime;
    const docBody = context.document.body;

    // search() returns an array of Ranges.
    const searchResults = docBody.search('the', { matchWholeWord: true });
    context.load(searchResults, 'items');
    await context.sync();

    // Record the system time.
    startTime = performance.now();

    for (var i = 0; i < searchResults.items.length; i++) {
      searchResults.items[i].font.highlightColor = '#FFFF00';

      await context.sync(); // SYNCHRONIZE IN EACH ITERATION
    }
    
    // await context.sync(); // SYNCHRONIZE AFTER THE LOOP

    // Record the system time again then calculate how long the operation took.
    endTime = performance.now();
    console.log("The operation took: " + (endTime - startTime) + " milliseconds.");
  })
}
```

O código anterior levou 1 segundo para concluir em um documento com 200 instâncias de "The" no Word no Windows. Mas quando a `await context.sync();` linha dentro do loop é comentada e a mesma linha após o loop é descomentado, a operação levou apenas 1/10 de um segundo. No Word na Web (com borda como o navegador), levou 3 segundos completos à sincronização dentro do loop e somente 6/décimos de segundo com a sincronização após o loop, cerca de cinco vezes mais velozes. Em um documento com 2000 instâncias de "a", ele levou (no Word na Web) 80 segundos com a sincronização dentro do loop e apenas 4 segundos com a sincronização após o loop, cerca de 20 vezes mais veloz.

> [!NOTE]
> Vale a pena perguntar se a versão de sincronização no loop será executada mais rapidamente se as sincronizações foram executadas simultaneamente, o que poderia ser feito simplesmente removendo a `await` palavra-chave da parte frontal do `context.sync()` . Isso fará com que o tempo de execução inicie a sincronização e, em seguida, inicie imediatamente a próxima iteração do loop sem aguardar a conclusão da sincronização. No entanto, isso não é uma boa solução como se fosse movimentar o `context.sync` loop completo por esses motivos:
>
> - Assim como os comandos em um trabalho em lotes de sincronização são enfileirados, os trabalhos em lotes em si são colocados na fila no Office, mas o Office não dá suporte a mais de 50 trabalhos em lote na fila. Mais erros de gatilho. Portanto, se houver mais de 50 iterações em um loop, haverá uma chance de que o tamanho da fila seja excedido. Quanto maior o número de iterações, maior a chance de isso acontecer. 
> - "Simultâneo" não significa simultaneamente. Ainda será mais demorado executar várias operações de sincronização do que executar uma.
> - Não é garantido que as operações simultâneas sejam concluídas na mesma ordem em que foram iniciadas. No exemplo anterior, não importa a ordem em que a palavra "a" é realçada, mas há situações em que é importante que os itens da coleção sejam processados na ordem.

## <a name="reading-values-from-the-document-with-the-split-loop-pattern"></a>Ler valores do documento com o padrão de loop dividido

`context.sync`A prevenção de s dentro de um loop se torna mais desafiador quando o código deve *ler* uma propriedade dos itens da coleção à medida que processa cada um. Suponha que seu código precise iterar todos os controles de conteúdo em um documento do Word e registre o texto do primeiro parágrafo associado a cada controle. Seus instintos de programação podem levar você a fazer um loop sobre os controles, carregar a `text` propriedade de cada parágrafo (primeiro), chamar o `context.sync` preenchimento do objeto de parágrafo de proxy com o texto do documento e, em seguida, fazê-lo. Apresentamos um exemplo a seguir.

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

Neste cenário, para evitar ter um `context.sync` loop, você deve usar um padrão que chamamos o padrão de **loop dividido** . Vamos ver um exemplo concreto do padrão antes de obtermos uma descrição formal dele. Veja como o padrão de loop dividido pode ser aplicado ao trecho de código anterior. Observe o seguinte sobre este código:

- Agora há dois loops e o `context.sync` vêm entre eles, portanto, não há nenhum `context.sync` loop.
- O primeiro loop itera através dos itens no objeto da coleção e carrega a `text` propriedade da mesma forma que o loop original, mas o primeiro loop não pode registrar o texto do parágrafo porque ele não contém mais um `context.sync` para popular a `text` Propriedade do `paragraph` objeto proxy. Em vez disso, ele adiciona o `paragraph` objeto a uma matriz.
- O segundo loop itera através da matriz que foi criada pelo primeiro loop e registra o `text` de cada `paragraph` Item. Isso é possível porque o `context.sync` que vem entre os dois loops preencheram todas as `text` Propriedades.

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

O exemplo anterior sugere o procedimento a seguir para transformar um loop que contenha um `context.sync` no padrão de loop dividido: 

1. Substitua o loop por dois loops.
2. Crie um primeiro loop para iterar sobre a coleção e adicione cada item a uma matriz enquanto também estiver carregando qualquer Propriedade do item que seu código precisa ler. 
3. Após o primeiro loop, chame `context.sync` para preencher os objetos de proxy com todas as propriedades carregadas. 
4. Siga as instruções `context.sync` com um segundo loop para iterar sobre a matriz criada no primeiro loop e ler as propriedades carregadas.

## <a name="processing-objects-in-the-document-with-the-correlated-objects-pattern"></a>Processamento de objetos no documento com o padrão de objetos correlacionados

Consideremos um cenário mais complexo em que o processamento dos itens na coleção requer dados que não estão nos próprios itens. O cenário prevê um suplemento do Word que opera em documentos criados a partir de um modelo com texto clichê. Dispersa no texto são uma ou mais instâncias das seguintes cadeias de caracteres de espaço reservado: "{Coordinator}", "{Deputy}" e "{Manager}". O suplemento substitui cada espaço reservado pelo nome de algumas pessoas. A interface do usuário do suplemento não é importante para este artigo. Por exemplo, ele pode ter um painel de tarefas com três caixas de texto, cada uma rotulada com um dos espaços reservados. O usuário insere um nome em cada caixa de texto e, em seguida, pressiona um botão **substituir** . O manipulador para o botão cria uma matriz que mapeia os nomes para os espaços reservados e, em seguida, substitui cada espaço reservado pelo nome atribuído. 

Você não precisa realmente produzir um suplemento com esta interface do usuário para experimentar o código. Você pode usar a [ferramenta de laboratório de script para criar](../overview/explore-with-script-lab.md) um protótipo do código importante. Use a instrução de atribuição a seguir para criar a matriz de mapeamento.

```javascript
const jobMapping = [
        { job: "{Coordinator}", person: "Sally" },
        { job: "{Deputy}", person: "Bob" },
        { job: "{Manager}", person: "Kim" }
    ];
```

O código a seguir mostra como você pode substituir cada espaço reservado por seu nome atribuído, se você usou `context.sync` dentro de loops.

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

No código anterior, há um loop externo e interno. Cada uma delas contém um `context.sync` . Com base no primeiro trecho de código deste artigo, provavelmente você verá que o `context.sync` loop interno pode ser movido após o loop interno. Mas isso ainda deixaria o código com um `context.sync` (dois deles na verdade) no loop externo. O código a seguir mostra como você pode remover `context.sync` dos loops. Discutimos o código a seguir.

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

Observação o código usa o padrão de loop dividido:

- O loop externo do exemplo anterior foi dividido em dois. (O segundo loop tem um loop interno, que é esperado porque o código está em iteração em um conjunto de trabalhos (ou espaços reservados) e dentro desse conjunto que está Iterando nos intervalos correspondentes.)
- Há um `context.sync` após cada loop principal, mas não `context.sync` dentro de nenhum loop. 
- O segundo loop principal itera por meio de uma matriz criada no primeiro loop.

Mas a matriz criada no primeiro loop *não contém apenas* um objeto do Office como o primeiro loop fazia na seção [lendo valores do documento com o padrão de loop dividido](#reading-values-from-the-document-with-the-split-loop-pattern). Isso ocorre porque algumas das informações necessárias para processar os objetos de intervalo do Word não estão nos próprios objetos Range, mas, em vez disso, vêm da `jobMapping` matriz. 

Portanto, os objetos na matriz criados no primeiro loop são objetos personalizados que têm duas propriedades. O primeiro é uma matriz de intervalos de palavras que correspondem a um título de trabalho específico (ou seja, uma cadeia de caracteres de espaço reservado) e a segunda é uma cadeia de caracteres que fornece o nome da pessoa atribuída ao trabalho. Isso torna o loop final fácil de escrever e fácil de ler, pois todas as informações necessárias para processar um determinado intervalo estão contidas no mesmo objeto personalizado que contém o intervalo. O nome que deve substituir _ **correlatedObject**. rangesMatchingJob. Items [j]_ é a outra Propriedade do mesmo objeto: _ **correlatedObject**. personAssignedToJob_. 

Chamamos essa variação do padrão de loop dividido do padrão de **objetos correlacionados** . A idéia geral é que o primeiro loop cria uma matriz de objetos personalizados. Cada objeto tem uma propriedade cujo valor é um dos itens em um objeto do conjunto do Office (ou uma matriz desses itens). O objeto personalizado tem outras propriedades, cada uma das quais fornece informações necessárias para processar os objetos do Office no loop final. Consulte a seção [outros exemplos desses padrões](#other-examples-of-these-patterns) para obter um link para um exemplo em que o objeto correlacionáe personalizado tem mais de duas propriedades.

Uma restrição adicional: às vezes, leva mais de um loop apenas para criar a matriz de objetos correlacionados personalizados. Isso pode acontecer se você precisar ler uma propriedade de cada membro de um objeto do conjunto do Office apenas para coletar informações que serão usadas para processar outro objeto de coleção. (Por exemplo, seu código precisa ler os títulos de todas as colunas em uma tabela do Excel porque seu suplemento aplicará um formato de número às células de algumas colunas com base no título dessa coluna.) Mas você sempre pode manter o `context.sync` s entre os loops, e não um loop. Consulte a seção [outros exemplos desses padrões](#other-examples-of-these-patterns) para obter um exemplo.

## <a name="other-examples-of-these-patterns"></a>Outros exemplos desses padrões

- Para obter um exemplo muito simples para o Excel que usa `Array.forEach` loops, consulte a resposta aceita para esta pilha de excedente: [é possível enfileirar mais de um contexto. Load antes de Context. Sync?](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)
- Para obter um exemplo simples para o Word que usa `Array.forEach` loops e não usa `async` / `await` a sintaxe, consulte a resposta aceita para esta pilha de excedente: [iterar sobre todos os parágrafos com controles de conteúdo com a API JavaScript do Office](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api).
- Para obter um exemplo do Word que está escrito em TypeScript, confira o suplemento do Word de exemplo [Angular2 de estilo](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker), especialmente o arquivo [Word. Document. Service. TS](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts). Ele tem uma mistura `for` e `Array.forEach` loops.
- Para obter um exemplo avançado de palavra, importe [essa essência](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab) para a [ferramenta de laboratório de scripts](../overview/explore-with-script-lab.md). Para o contexto no uso da essência, confira a resposta aceita para o documento de pergunta de estouro de pilha [não sincronizado após substituir o texto](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text). Este exemplo cria um tipo de objeto correlacionado personalizado que tem três propriedades. Ele usa um total de três loops para construir a matriz de objetos correlacionados e dois loops para fazer o processamento final. Há uma mistura de `for` e `Array.forEach` loops.
- Embora não seja estritamente um exemplo dos padrões de loop de divisão ou objetos correlacionados, há um exemplo avançado do Excel que mostra como converter um conjunto de valores de célula em outras moedas com apenas um único `context.sync` . Para experimentá-lo, abra a [ferramenta de laboratório de script](../overview/explore-with-script-lab.md) e navegue até o exemplo de conversor de **moeda** . 

## <a name="when-should-you-not-use-the-patterns-in-this-article"></a>Quando você *não* deve usar os padrões neste artigo?

O Excel não pode ler mais de 5 MB de dados em uma determinada chamada de `context.sync` . Se esse limite for excedido, um erro será gerado. (Para obter mais informações, consulte [limites de transferência de dados do Excel](../develop/common-coding-issues.md#excel-data-transfer-limits).) É muito raro que esse limite seja abordado, mas se houver uma chance de que isso aconteça com seu suplemento, o código *não* deverá carregar todos os dados em um único loop e seguir o loop com um `context.sync` . Mas você ainda deve evitar ter um `context.sync` em cada iteração de um loop sobre um objeto de coleção. Em vez disso, defina subconjuntos dos itens na coleção e execute o loop sobre cada subconjunto por vez, com um `context.sync` entre os loops. Você pode estruturar isso com um loop externo que se repete nos subconjuntos e contém o `context.sync` em cada uma dessas iterações externas.
