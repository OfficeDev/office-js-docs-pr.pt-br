---
ms.date: 01/08/2020
description: Criar uma função personalizada no Excel para o Suplemento do Office.
title: Criar funções personalizadas no Excel
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 97037f201a237cdc6dae551552a0a1609a58b34c
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839870"
---
# <a name="create-custom-functions-in-excel"></a>Criar funções personalizadas no Excel

Funções personalizadas permitem que desenvolvedores adicionem novas funções do Excel definindo essas funções em JavaScript como parte de um suplemento. Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa no Excel, como `SUM()`.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

A imagem animada a seguir mostra a sua pasta de trabalho solicitando uma função que você criou com o JavaScript ou o Typescript. Neste exemplo, a função personalizada `=MYFUNCTION.SPHEREVOLUME` calcula o volume de uma esfera.

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

O código a seguir define a função personalizada `=MYFUNCTION.SPHEREVOLUME`.

```js
/**
 * Returns the volume of a sphere.
 * @customfunction
 * @param {number} radius
 */
function sphereVolume(radius) {
  return Math.pow(radius, 3) * 4 * Math.PI / 3;
}
```

> [!TIP]
> Se seu suplemento de função personalizada usará um painel de tarefas ou um botão da faixa de opções, além de executar o código de função personalizada, você precisará configurar um tempo de execução de JavaScript compartilhado. Consulte [Configure seu Suplemento do Office para usar em um tempo de execução do JavaScript compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md) para saber mais.

## <a name="how-a-custom-function-is-defined-in-code"></a>Como uma função personalizada é definida em código

Se você usar o [Gerador Yo Office](https://github.com/OfficeDev/generator-office) para criar um projeto de suplemento funções personalizadas do Excel, ele criará os arquivos que controlam as funções e o painel de tarefas. Vamos nos concentrar em arquivos que são importantes para funções personalizadas:

| File | Formato de arquivo | Descrição |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>ou<br/>**./src/functions/functions.ts** | JavaScript<br/>ou<br/>TypeScript | Contém o código que define funções personalizadas. |
| **./src/functions/functions.html** | HTML | Fornece uma referência&lt;script&gt;ao arquivo JavaScript que define funções personalizadas. |
| **./manifest.xml** | XML | Especifica o local de vários arquivos que a sua função personalizada usa, como as funções personalizadas JavaScript, JSON e arquivos HTML. Ele também lista os locais de arquivos do painel de tarefas, os arquivos de comando e especifica o tempo de execução que suas funções personalizadas devem usar. |

### <a name="script-file"></a>Arquivo de script

O arquivo de script (**./src/functions/functions.js** ou **./src/functions/functions.ts**) contém o código que define funções e comentários que definem a função.

O código a seguir define a função personalizada `add`. Os comentários do código são usados para gerar um arquivo de metadados JSON que descreve a função personalizada ao Excel. O necessário `@customfunction` comentário é declarado primeiro, para indicar que se trata de uma função personalizada. Em seguida, dois parâmetros são declarados, `first` e `second`, seguidos por suas propriedades de `description`. Por fim, uma `returns` descrição é fornecida. Para obter mais informações sobre quais comentários são necessários para sua função personalizada, confira [Gerar automaticamente os metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md).

```js
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number.
 * @param second Second number.
 * @returns The sum of the two numbers.
 */

function add(first, second){
  return first + second;
}
```

### <a name="manifest-file"></a>Arquivo de manifesto

O arquivo de manifesto XML para um suplemento que define funções personalizadas (**./manifest.xml** no projeto que o gerador de Yo Office cria) faz várias coisas:

- Define o espaço de nomes das suas funções personalizadas. Um namespace se direciona para suas funções personalizadas para ajudar os clientes a identificar suas funções como parte do seu suplemento.
- Usa os elementos `<ExtensionPoint>` e `<Resources>` que são exclusivos de um manifesto de funções personalizadas. Esses elementos contêm informações sobre os locais dos arquivos JavaScript, JSON e HTML.
- Especifica o tempo de execução a ser usado para a sua função personalizada. Recomendamos sempre usar um tempo de execução compartilhado, a menos que você tenha uma necessidade específica para outro tempo de execução, porque um tempo de execução compartilhado permite o compartilhamento de dados entre funções e o painel de tarefas. Observe que usar um tempo de execução compartilhado significa que seu suplemento usará o Internet Explorer 11, não o Microsoft Edge.

Se você estiver usando o gerador do Yo Office para criar arquivos, recomendamos ajustar o manifesto para usar o tempo de execução compartilhado, uma vez que esse não é o padrão para esses arquivos. Para alterar o manifesto, siga as instruções no [Configurar seu suplemento do Excel para usar um de tempo de execução JavaScript compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

Para ver um manifesto funcional completo de um suplemento de amostra, consulte [esse repositório do GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml).

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="coauthoring"></a>Coautoria

O Excel na Web e o Windows conectado a uma assinatura do Microsoft 365 permitem que você se conecte ao Excel. Se a pasta de trabalho usa uma função personalizada, seu colega será solicitado a carregar o suplemento da função personalizada. Depois de carregarem o suplemento, a função personalizada compartilhará resultados por meio de coautoria.

Para saber mais sobre coautoria, confira o tópico [Sobre o recurso de coautoria no Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).

## <a name="next-steps"></a>Próximas etapas

Quer experimentar funções personalizadas? Confira o simples [início rápido das funções personalizadas](../quickstarts/excel-custom-functions-quickstart.md) ou o mais detalhado [tutorial de funções personalizadas](../tutorials/excel-tutorial-create-custom-functions.md), caso ainda não tenha.

Outra maneira fácil de experimentar as funções personalizadas é usar o [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), que é um suplemento que permite com que você experimente as funções personalizadas diretamente no Excel. Você pode experimentar criar a sua própria função personalizada ou usar os exemplos disponíveis.

## <a name="see-also"></a>Confira também 
* [Saiba mais sobre o Programa para Desenvolvedores do Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
* [Conjuntos de requisitos de funções personalizadas](custom-functions-requirement-sets.md)
* [Diretrizes de nomenclatura de funções personalizadas](custom-functions-naming.md)
* [Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário](make-custom-functions-compatible-with-xll-udf.md)
* [Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
