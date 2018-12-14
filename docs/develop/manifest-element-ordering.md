---
title: Como encontrar a ordem correta dos elementos do manifesto
description: Saiba como encontrar a ordem correta na qual colocar elementos filho em um elemento pai.
ms.date: 11/16/2018
ms.openlocfilehash: 3efc95926b7562b0e68bbb6f4b13c47cc4ae6824
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270611"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a>Como encontrar a ordem correta dos elementos do manifesto

Os elementos XML do manifesto de um Suplemento do Office devem estar no elemento pai apropriado *e* em uma ordem específica em relação uns aos outros.

A ordem exigida é especificada nos arquivos XSD, na pasta [Esquemas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas). Os arquivos XSD são categorizados em subpastas para suplementos de painel de tarefas, conteúdo e email.

Por exemplo, no elemento `<OfficeApp>`, os elementos `<Id>`, `<Version>` e `<ProviderName>` devem aparecer nessa ordem. Se adicionar um elemento `<AlternateId>`, deverá colocá-lo entre os elementos `<Id>` e `<Version>`. Se algum dos elementos estiver na posição incorreta, o manifesto não será válido e o suplemento não será carregado.

> [!NOTE]
> O [Validador de Suplemento do Office](/office/dev/add-ins/testing/troubleshoot-manifest#validate-your-manifest-with-the-office-add-in-validator) usa a mesma mensagem de erro quando um elemento está fora de ordem, como ocorre quando um elemento está no pai incorreto. A mensagem de erro informa que o elemento não é um elemento filho válido do elemento pai. Caso receba este erro, mas a documentação de referência do elemento filho indique que ele *está* válido para o pai, talvez o problema seja o filho ter sido colocado na ordem incorreta.

Para encontrar a ordem correta dos elementos filho de um determinado elemento pai, faça os procedimentos a seguir. Este é um processo simplificado porque os arquivos XSD são bastante complexos. A análise completa dos arquivos XSD está fora do escopo deste documento.

1. Abra a subpasta em [Esquemas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) para o tipo de suplemento que você está criando. 
2. Abra o arquivo XSD no qual o elemento pai é definido como um tipo complexo. Se você não souber qual arquivo tem a definição, talvez seja necessário realizar a etapa 3 em vários arquivos até encontrá-la.
3. Procure `<xs:complexType name="PARENT_ELEMENT">`, em que PARENT_ELEMENT é o nome do elemento pai.
4. Dentro da definição de PARENT_ELEMENT, normalmente há um elemento chamado `<xs:sequence>`. Veja a seguir a definição de `<SuperTip>` em [TaskPaneAppVersionOverridesV1_0.xsd](https://raw.githubusercontent.com/OfficeDev/office-js-docs-pr/master/docs/overview/schemas/taskpane/TaskPaneAppVersionOverridesV1_0.xsd).

```xml
  <xs:complexType name="Supertip">
    <xs:annotation>
      <xs:documentation>
        Specifies the super tip for this control.
      </xs:documentation>
    </xs:annotation>
    <xs:sequence>
      <xs:element name="Title" type="bt:ShortResourceReference" minOccurs="1" maxOccurs="1" />
      <xs:element name="Description" type="bt:LongResourceReference" minOccurs="1" maxOccurs="1" />
    </xs:sequence>
  </xs:complexType>
```

A `<xs:sequence>` descreve os possíveis elementos filho, *na ordem em que devem aparecer*. Isto *não* significa que todos eles sejam obrigatórios. Se o valor `minOccurs` de um elemento filho for **0**, então esse elemento será opcional. *Mas se ele estiver presente, deverá estar na ordem especificada pelo elemento `<xs:sequence>`*.

Se não houver um elemento `<xs:sequence>` ou *se houver*, mas o elemento filho não estiver relacionado, mesmo que a respectiva documentação de referência indique que ele *esteja* válido para o pai, então a definição de tipo complexo do elemento pai terá sido estendida com elementos filho adicionais em outro local no arquivo XSD. Por exemplo, a definição do tipo complexo `OfficeApp` não relaciona `Requirements` como um possível filho. Mas posteriormente no arquivo, dentro da definição do tipo complexo `TaskPaneApp`, a definição de `OfficeApp` será estendida e `Requirements` será adicionado como um elemento filho adicional válido.

Para encontrar as definições estendidas, faça o seguinte:

1. No início do arquivo, procure `<xs:extension base="PARENT_ELEMENT">`, em que PARENT_ELEMENT é o nome do elemento pai. Talvez haja mais de uma extensão.
2. Procure a extensão que seja relevante para o contexto no qual você está trabalhando. Por exemplo, o tipo complexo `OfficeApp` está estendido dentro dos tipos complexos `ContentApp` e `MailApp`, bem como dentro do tipo complexo `TaskPaneApp`.

Cada `<xs:extension base="PARENT_ELEMENT">` no arquivo tem o próprio `<xs:sequence>` que relaciona elementos filho adicionais válidos para o pai. Os elementos filho de uma lista estendida devem ser colocados sempre *após* os elementos filho da lista original, na definição de tipo complexo do pai.

## <a name="see-also"></a>Confira também

- [Referência de esquema para manifestos de Suplementos do Office (versão 1.1)](../develop/add-in-manifests.md)
