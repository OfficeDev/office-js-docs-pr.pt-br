---
ms.date: 04/18/2019
description: Solução de problemas comuns em funções personalizadas do Excel.
title: Solução de problemas de funções personalizadas (versão prévia)
localization_priority: Priority
ms.openlocfilehash: cf54aa3b719b7893799df5d1c5206c6fb904be69
ms.sourcegitcommit: 44c61926d35809152cbd48f7b97feb694c7fa3de
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/22/2019
ms.locfileid: "31959101"
---
# <a name="troubleshoot-custom-functions"></a>Solução de problemas de funções personalizadas

Ao desenvolver funções personalizadas, você poderá encontrar erros no produto durante a criação e testes das funções.

Para resolver problemas, você pode [habilitar o log de tempo de execução para capturar erros](#enable-runtime-logging) e consultar as [mensagens de erro nativas do Excel](#check-for-excel-error-messages). Além disso, verifique se há erros comuns, como não [verificar certificados ssl](#my-add-in-wont-load-verify-certificates) de forma adequada, [deixar promessas não resolvidas](#ensure-promises-return) e esquecer de [associar as funções](#my-functions-wont-load-associate-functions).

## <a name="enable-runtime-logging"></a>Habilitar o log de tempo de execução

Se estiver testando o suplemento do Office no Windows, você deverá [habilitar o log de tempo de execução](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in). O log de tempo de execução entrega instruções `console.log` a um arquivo de log separado criado para ajudar você a descobrir problemas. As instruções abrangem vários erros, incluindo os relacionados ao arquivo de manifesto XML do suplemento, condições do tempo de execução ou a instalação de funções personalizadas.  Saiba mais sobre o log de tempo de execução em [Usar o log de tempo de execução para depurar seu suplemento](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).  

### <a name="check-for-excel-error-messages"></a>Verificar se há mensagens de erro do Excel

O Excel tem diversas mensagens de erro internas que serão retornadas para uma célula se houver um erro de cálculo. As funções personalizadas usam apenas as seguintes mensagens de erro: `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A` e `#BUSY!`.

## <a name="common-issues"></a>Problemas comuns

### <a name="my-add-in-wont-load-verify-certificates"></a>Meu suplemento não carrega: verificar certificados

Se o suplemento não for devidamente instalado, verifique se os certificados SSL estão configurados corretamente para o servidor Web que hospeda seu suplemento. Normalmente, se houver um problema com os certificados SSL, você verá uma mensagem de erro no Excel avisando que não foi possível instalar seu suplemento corretamente. Saiba mais em [Adicionar certificados autoassinados como certificado raiz de confiança](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).

### <a name="my-functions-wont-load-associate-functions"></a>Minhas funções não carregam: associar funções

No arquivo de script das funções personalizadas, você precisa associar cada função personalizada à respectiva ID especificada no [arquivo de metadados JSON](custom-functions-json.md). Isso é feito usando o método `CustomFunctions.associate()`. Normalmente, essa chamada de método é feita após cada função ou no final do arquivo de script. Se uma função personalizada não estiver associada, ele não funcionará.

O exemplo a seguir mostra uma função add, seguida pelo nome `add` da função que está sendo associada a `ADD` da id JSON correspondente.

```js
function add(first, second){
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

Saiba mais sobre esse processo em [Associar os nomes de função com metadados JSON](/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata).

### <a name="cant-open-add-in-from-localhost-use-a-local-loopback-exception"></a>Não é possível abrir um suplemento de um localhost: utilize uma exceção de loopback local

Se você vir o erro "Não é possível abrir este suplemento de um localhost", será necessário habilitar uma exceção de loopback local. Para obter detalhes sobre como fazer isso, confira [este artigo de suporte da Microsoft](https://support.microsoft.com/pt-BR/help/4490419/local-loopback-exemption-does-not-work).

### <a name="ensure-promises-return"></a>Garantir que as promessas retornem resultados

Quando o Excel está aguardando a conclusão de uma função personalizada, ele exibe #BUSY! na célula. Se o código da função personalizada retornar uma promessa, mas a promessa não retornar um resultado, o Excel continuará exibindo #BUSY!. Verifique suas funções para garantir que as promessas estejam retornando corretamente um resultado para uma célula.

## <a name="reporting-feedback"></a>Fornecer comentários

Se você tiver problemas que não estão descritos aqui, fale conosco. Há duas maneiras de relatar problemas.

### <a name="in-excel-on-windows-or-mac"></a>No Excel para Windows ou Mac

Se estiver usando o Excel para Windows ou Mac, envie comentários à equipe de extensibilidade do Office diretamente do Excel. Para fazer isso, selecione **Arquivo -> Comentários -> Enviar um Rosto Triste**. Enviando um Rosto Triste, você fornece os registros necessários para entendermos o problema que você está enfrentando.

### <a name="in-github"></a>No Github

Sinta-se à vontade para enviar problemas encontrados através do recurso "Comentários do conteúdo" na parte inferior de todas as páginas de documentação ou [informe um novo problema diretamente no repositório de funções personalizadas](https://github.com/OfficeDev/Excel-Custom-Functions/issues).

## <a name="see-also"></a>Confira também

* [Metadados de funções personalizadas](custom-functions-json.md)
* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
* [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).
* [Log de alteração de funções personalizadas](custom-functions-changelog.md)
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
