---
title: Evite usar o método context.sync em loops
description: Saiba como usar o loop dividido e os padrões de objetos correlacionados para evitar chamar context.sync em um loop.
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: 64cfd5cd350746ba07e1a98986a4bd7811431475
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349137"
---
# <a name="avoid-using-the-contextsync-method-in-loops"></a>Evite usar o método context.sync em loops

> [!NOTE]
> Este artigo supõe que você esteja além do estágio inicial de trabalhar com pelo menos uma das quatro APIs JavaScript específicas do Office aplicativo para Excel, Word, OneNote e Visio que usam um sistema em lotes para interagir com o &mdash; &mdash; documento Office. Em particular, você deve saber o que uma chamada `context.sync` faz e deve saber o que é um objeto de coleção. Se você não estiver nesse estágio, comece com Understanding [the Office API JavaScript](../develop/understanding-the-javascript-api-for-office.md) e a documentação vinculada em "específico do aplicativo" nesse artigo.

Para alguns cenários de programação em Office Add-ins que usam um dos modelos de API específicos do aplicativo (para Excel, Word, OneNote e Visio), seu código precisa ler, gravar ou processar alguma propriedade de cada membro de um objeto de coleção. Por exemplo, um Excel que precisa obter os valores de cada célula em uma coluna de tabela específica ou um complemento do Word que precisa realçar cada instância de uma cadeia de caracteres no documento. Você precisa iterar sobre os membros na propriedade do objeto da coleção; mas, por motivos de desempenho, você precisa evitar chamar em cada `items` `context.sync` iteração do loop. Cada chamada de é uma viagem de ida e volta do `context.sync` add-in para o Office documento. Viagens de ida e volta repetidas prejudicam o desempenho, especialmente se o add-in estiver em execução Office na Web porque as idas e voltas vão pela Internet.

> [!NOTE]
> Todos os exemplos neste artigo usam loops, mas as práticas descritas se aplicam a qualquer instrução de loop que possa iterar através de uma `for` matriz, incluindo o seguinte:
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

## <a name="writing-to-the-document"></a>Escrevendo no documento

No caso mais simples, você só está escrevendo para membros de um objeto de coleção, não lendo suas propriedades. Por exemplo, o código a seguir realça em amarelo todas as instâncias de "o" em um documento do Word.

> [!NOTE]
> Geralmente, é uma boa prática colocar ter uma final pouco antes do caractere "}" de fechamento do método de aplicativo `context.sync` `run` (como `Excel.run` , , `Word.run` etc.). Isso acontece porque o método faz uma chamada oculta como a última coisa que ele faz se, e somente se, há comandos em fila que ainda não `run` `context.sync` foram sincronizados. O fato de essa chamada estar oculta pode ser confuso, portanto, geralmente recomendamos que você adicione o `context.sync` explícito . No entanto, considerando que este artigo se trata de minimizar chamadas de , na verdade, é mais confuso adicionar um `context.sync` final totalmente `context.sync` desnecessário. Portanto, neste artigo, o deixamos de fora quando não há comandos não sincronizados no final do `run` .

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

O código anterior levou 1 segundo completo para ser concluído em um documento com 200 instâncias de "o" no Word no Windows. Mas quando a linha dentro do loop é comentada para fora e a mesma linha logo após o loop ser descomentado, a operação levou apenas um décimo de `await context.sync();` segundo. No Word na Web (com Edge como navegador), levou 3 segundos completos com a sincronização dentro do loop e apenas 6/10ths de um segundo com a sincronização após o loop, cerca de cinco vezes mais rápido. Em um documento com 2.000 instâncias de "the", ele levou (em Word na Web) 80 segundos com a sincronização dentro do loop e apenas 4 segundos com a sincronização após o loop, cerca de 20 vezes mais rápido.

> [!NOTE]
> Vale a pena perguntar se a versão de sincronização dentro do loop seria executada mais rapidamente se as sincronizações fosse executadas simultaneamente, o que poderia ser feito simplesmente removendo a palavra-chave da frente do `await` `context.sync()` . Isso faria com que o tempo de execução iniciasse a sincronização e iniciasse imediatamente a próxima iteração do loop sem aguardar a conclusão da sincronização. No entanto, essa solução não é tão boa quanto mover completamente o `context.sync` loop por esses motivos:
>
> - Assim como os comandos em um trabalho em lotes de sincronização são enfilados, os trabalhos em lotes em si são enfilados no Office, mas o Office dá suporte a não mais de 50 trabalhos em lotes na fila. Mais erros disparam. Portanto, se houver mais de 50 iterações em um loop, há uma chance de que o tamanho da fila seja excedido. Quanto maior o número de iterações, maior será a chance de isso acontecer. 
> - "Simultaneamente" não significa simultaneamente. Ainda levaria mais tempo para executar várias operações de sincronização do que executar uma.
> - Operações simultâneas não são garantidas na mesma ordem em que iniciaram. No exemplo anterior, não importa qual ordem a palavra "the" é realçada, mas há cenários em que é importante que os itens da coleção sejam processados em ordem.

## <a name="reading-values-from-the-document-with-the-split-loop-pattern"></a>Lendo valores do documento com o padrão de loop dividido

Evitar s dentro de um loop se torna mais desafiador quando o código deve ler uma propriedade dos itens da coleção à `context.sync` medida que processa cada um.  Suponha que seu código precise iterar todos os controles de conteúdo em um documento do Word e registrar o texto do primeiro parágrafo associado a cada controle. Seus instintos de programação podem levar você a fazer um loop sobre os controles, carregar a propriedade de cada parágrafo (primeiro), chamar para preencher o objeto de parágrafo proxy com o texto do documento e, em `text` `context.sync` seguida, registrá-lo. Apresentamos um exemplo a seguir.

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

Nesse cenário, para evitar ter um loop, você deve usar um padrão que chamamos de padrão de `context.sync` loop **dividido.** Vamos ver um exemplo concreto do padrão antes de chegarmos a uma descrição formal dele. Veja como o padrão de loop dividido pode ser aplicado ao trecho de código anterior. Observe o seguinte sobre este código.

- Agora há dois loops e o `context.sync` entre eles, portanto, não há `context.sync` nenhum loop dentro de ambos.
- O primeiro loop itera pelos itens no objeto da coleção e carrega a propriedade da mesma forma que o loop original, mas o primeiro loop não pode registrar o texto do parágrafo porque ele não contém mais um para preencher a propriedade do objeto `text` `context.sync` `text` `paragraph` proxy. Em vez disso, ele adiciona `paragraph` o objeto a uma matriz.
- O segundo loop itera pela matriz criada pelo primeiro loop e registra `text` o de cada `paragraph` item. Isso é possível porque o `context.sync` que veio entre os dois loops preencheu todas as `text` propriedades.

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

O exemplo anterior sugere o procedimento a seguir para transformar um loop que contém um `context.sync` no padrão de loop dividido.

1. Substitua o loop por dois loops.
2. Crie um primeiro loop para iterar sobre a coleção e adicione cada item a uma matriz enquanto também carrega qualquer propriedade do item que seu código precisa ler.
3. Após o primeiro loop, chame `context.sync` para preencher os objetos proxy com quaisquer propriedades carregadas.
4. Siga o `context.sync` com um segundo loop para iterar sobre a matriz criada no primeiro loop e ler as propriedades carregadas.

## <a name="processing-objects-in-the-document-with-the-correlated-objects-pattern"></a>Processar objetos no documento com o padrão de objetos correlacionados

Vamos considerar um cenário mais complexo em que o processamento dos itens na coleção exige dados que não estão nos itens em si. O cenário visualiza um complemento do Word que opera em documentos criados a partir de um modelo com algum texto clichê. Espalhados no texto estão uma ou mais instâncias das seguintes cadeias de caracteres de espaço reservado: "{Coordenador}", "{Deputy}" e "{Manager}". O complemento substitui cada espaço reservado pelo nome de uma pessoa. A interface do usuário do complemento não é importante para este artigo. Por exemplo, ele poderia ter um painel de tarefas com três caixas de texto, cada uma rotulada com um dos espaço reservados. O usuário insra um nome em cada caixa de texto e pressiona um **botão Substituir.** O manipulador do botão cria uma matriz que mapeia os nomes para os espaço reservados e substitui cada espaço reservado pelo nome atribuído. 

Você não precisa realmente produzir um complemento com essa interface do usuário para experimentar o código. Você pode usar a ferramenta [Script Lab para](../overview/explore-with-script-lab.md) protótipo do código importante. Use a instrução de atribuição a seguir para criar a matriz de mapeamento.

```javascript
const jobMapping = [
        { job: "{Coordinator}", person: "Sally" },
        { job: "{Deputy}", person: "Bob" },
        { job: "{Manager}", person: "Kim" }
    ];
```

O código a seguir mostra como você pode substituir cada espaço reservado por seu nome atribuído se você usou `context.sync` loops dentro.

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

No código anterior, há um loop externo e interno. Cada um deles contém um `context.sync` . Com base no primeiro trecho de código deste artigo, você provavelmente verá que o loop interno pode simplesmente ser movido `context.sync` após o loop interno. Mas isso ainda deixaria o código com `context.sync` um (dois deles, na verdade) no loop externo. O código a seguir mostra como você pode `context.sync` remover dos loops. Discutiremos o código abaixo.

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

Observe que o código usa o padrão de loop dividido:

- O loop externo do exemplo anterior foi dividido em dois. (O segundo loop tem um loop interno, que é esperado porque o código está iterando sobre um conjunto de trabalhos (ou espaço reservados) e dentro desse conjunto ele está iterando sobre os intervalos correspondentes.)
- Há um após `context.sync` cada loop principal, mas `context.sync` não dentro de qualquer loop.
- O segundo loop principal itera através de uma matriz criada no primeiro loop.

Mas a matriz criada no  primeiro loop não contém apenas um objeto Office como o primeiro loop fez na seção Ler valores do documento com o padrão de [loop dividido.](#reading-values-from-the-document-with-the-split-loop-pattern) Isso acontece porque algumas das informações necessárias para processar os objetos Intervalo do Word não estão nos objetos Range em si, mas vêm da `jobMapping` matriz.

Portanto, os objetos na matriz criada no primeiro loop são objetos personalizados que têm duas propriedades. O primeiro é uma matriz de Intervalos do Word que combinam com um título de trabalho específico (ou seja, uma cadeia de caracteres de espaço reservado) e a segunda é uma cadeia de caracteres que fornece o nome da pessoa atribuída ao trabalho. Isso torna o loop final fácil de gravar e fácil de ler porque todas as informações necessárias para processar um determinado intervalo estão contidas no mesmo objeto personalizado que contém o intervalo. O nome que deve substituir _**correlacionadoObject**.rangesMatchingJob.items[j]_ é a outra propriedade do mesmo objeto: _**correlatedObject**.personAssignedToJob_.

Chamamos essa variação do padrão de loop dividido do **padrão de objetos correlacionados.** A ideia geral é que o primeiro loop cria uma matriz de objetos personalizados. Cada objeto tem uma propriedade cujo valor é um dos itens em um objeto Office coleção Office (ou uma matriz desses itens). O objeto personalizado tem outras propriedades, cada uma delas fornece informações necessárias para processar os objetos Office no loop final. Consulte a seção [Outros exemplos desses](#other-examples-of-these-patterns) padrões para um link para um exemplo em que o objeto de correlação personalizado tem mais de duas propriedades.

Uma outra advertência: às vezes, é necessário mais de um loop apenas para criar a matriz de objetos de correlação personalizados. Isso pode acontecer se você precisar ler uma propriedade de cada membro de um objeto da coleção Office apenas para coletar informações que serão usadas para processar outro objeto de coleção. (Por exemplo, seu código precisa ler os títulos de todas as colunas em uma tabela Excel porque o seu complemento aplicará um formato numérico às células de algumas colunas com base no título dessa coluna.) Mas você sempre pode manter `context.sync` o s entre os loops, em vez de em um loop. Consulte a seção [Outros exemplos desses padrões](#other-examples-of-these-patterns) para um exemplo.

## <a name="other-examples-of-these-patterns"></a>Outros exemplos desses padrões

- Para obter um exemplo muito simples para Excel que usa loops, consulte a resposta aceita para esta pergunta Stack Overflow: É possível enfileir mais de um `Array.forEach` [context.load antes de context.sync?](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)
- Para um exemplo simples para o Word que usa loops e não usa sintaxe, consulte a resposta aceita para esta pergunta `Array.forEach` `async` / `await` Stack Overflow: [Iterating over](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api)all paragraphs with content controls with Office JavaScript API .
- Para ver um exemplo do Word escrito em TypeScript, consulte o exemplo [do Word Add-in Angular2 Style Checker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker), especialmente o arquivo [word.document.service.ts](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts). Ele tem uma mistura de `for` `Array.forEach` e loops.
- Para um exemplo avançado do Word, importe [esse gist](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab) para a [Script Lab ferramenta](../overview/explore-with-script-lab.md). Para contexto ao usar o gist, consulte a resposta aceita para a pergunta Stack Overflow [Document not in sync after replace text](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text). Este exemplo cria um tipo de objeto de correlação personalizado que tem três propriedades. Ele usa um total de três loops para construir a matriz de objetos correlacionados e mais dois loops para fazer o processamento final. Há uma mistura de `for` e `Array.forEach` loops.
- Embora não seja estritamente um exemplo do loop dividido ou dos padrões de objetos correlacionados, há uma amostra de Excel avançada que mostra como converter um conjunto de valores de célula em outras moedas com apenas um `context.sync` único . Para experimentar, abra a ferramenta [Script Lab e](../overview/explore-with-script-lab.md) navegue até o exemplo **Conversor de Moedas Conversor.**

## <a name="when-should-you-not-use-the-patterns-in-this-article"></a>Quando você *não deve* usar os padrões neste artigo?

Excel pode ler mais de 5 MB de dados em uma determinada chamada de `context.sync` . Se esse limite for excedido, será lançado um erro. (Consulte a seção "Excel de Excel" dos limites de recursos e otimização de desempenho para Office [de Office para](resource-limits-and-performance-optimization.md#excel-add-ins) obter mais informações.) É muito raro que esse limite seja abordado, mas se houver uma chance de isso acontecer  com o seu complemento, seu código não deve carregar todos os dados em um único loop e seguir o loop com `context.sync` um . Mas você ainda deve evitar ter uma `context.sync` em cada iteração de um loop sobre um objeto de coleção. Em vez disso, defina subconjuntos dos itens na coleção e loop sobre cada subconjunto, por sua vez, com um `context.sync` entre os loops. Você pode estruturar isso com um loop externo que itera sobre os subconjuntos e contém o em cada `context.sync` uma dessas iterações externas.
