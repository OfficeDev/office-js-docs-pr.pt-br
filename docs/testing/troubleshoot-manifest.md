---
title: Validar o manifesto de suplemento do Office
description: Saiba como validar o manifesto de um Office add-in usando o esquema XML e outras ferramentas.
ms.date: 09/18/2020
localization_priority: Normal
ms.openlocfilehash: ae2254bcd14b70b7bbd6c71dbbac2c90f53dd2580115a4a532931cee8d674c32
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57097806"
---
# <a name="validate-an-office-add-ins-manifest"></a>Validar o manifesto de suplemento do Office

Talvez você queira validar o arquivo de manifesto do seu suplemento para garantir que ele está correto e completo. A validação também pode identificar problemas que estejam causando o erro "seu manifesto de suplemento não é válido" quando você tenta realizar o sideload do seu suplemento. Este artigo descreve várias maneiras de validar o arquivo de manifesto.

> [!NOTE]
> Para saber mais sobre como usar o log de tempo de execução para solucionar problemas no manifesto de suplemento, confira [Depurar seu suplemento com o log de tempo de execução](runtime-logging.md).

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a>Validar o manifesto com o gerador Yeoman para Suplementos do Office

Se você usou o [gerador de Yeoman para suplementos](https://www.npmjs.com/package/generator-office) do Office para criar seu suplemento, você também pode usá-lo para validar o arquivo de manifesto do seu projeto. Execute o seguinte comando no diretório raiz do seu projeto.

```command&nbsp;line
npm run validate
```

![GIF animado que mostra o validador Yo Office sendo executado na linha de comando e gerando resultados que mostram Validation Passed.](../images/yo-office-validator.gif)

> [!NOTE]
> Para ter acesso a essa funcionalidade, o projeto de suplemento deve ter sido criado usando o [Gerador Yeoman para Suplementos do Office](https://www.npmjs.com/package/generator-office) versão 1.1.17 ou posterior.

## <a name="validate-your-manifest-with-office-addin-manifest"></a>Valide seu manifesto com o office-addin-manifest

Se você não tiver usado o [gerador Yeoman para Suplementos do Office](https://www.npmjs.com/package/generator-office) para criar seu suplemento, você também pode usá-lo para validar o arquivo de manifesto usando o[office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).

1. Instale o [Node.js](https://nodejs.org/download/).

1. Abra um prompt de comando e instale o validador com o seguinte comando.

    ```command&nbsp;line
    npm install -g office-addin-manifest
    ```

1. Execute o seguinte comando *no diretório raiz do seu projeto*.

    ```command&nbsp;line
    npm run validate
    ```

    > [!NOTE]
    > Se esse comando não estiver disponível ou não estiver funcionando, execute o seguinte comando para forçar o uso da versão mais recente da ferramenta office-addin-manifest (substituindo pelo nome do arquivo `MANIFEST_FILE` de manifesto).
    >
    > ```command&nbsp;line
    > npx --ignore-existing office-addin-manifest validate MANIFEST_FILE
    > ```

## <a name="validate-your-manifest-against-the-xml-schema"></a>Validar seu manifesto em relação ao esquema XML

É possível validar um manifesto em relação aos arquivos de [Definição de esquema XML (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8). Para ajudar a garantir que o arquivo de manifesto segue o esquema correto, incluindo todos os namespaces para os elementos que você está usando. Se você copiou elementos de outros manifestos da amostra, verifique se também **incluiu os namespaces apropriados**. É possível usar uma ferramenta de validação de esquema XML para executar essa validação.

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a>Para usar uma ferramenta de validação de esquema XML da linha de comando para validar seu manifesto

1. Instale o [tar](https://www.gnu.org/software/tar/) e o [libxml](http://xmlsoft.org/FAQ.html), caso ainda não tenha o feito.

1. Execute o comando a seguir. Substitua `XSD_FILE` pelo caminho para o arquivo XSD do manifesto e `XML_FILE` pelo caminho para o arquivo XML do manifesto.

    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="see-also"></a>Confira também

- [Manifesto XML dos Suplementos do Office](../develop/add-in-manifests.md)
- [Limpar o cache do Office](clear-cache.md)
- [Depurar seu suplemento com o log de tempo de execução](runtime-logging.md)
- [Realizar sideload de suplementos do Office para teste](sideload-office-add-ins-for-testing.md)
- [Depurar suplementos do Office](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
