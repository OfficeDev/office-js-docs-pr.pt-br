---
ms.date: 05/03/2019
description: Solução de problemas comuns em funções personalizadas do Excel.
title: Solução de problemas das funções personalizadas
localization_priority: Priority
ms.openlocfilehash: 04da6d58c2610130961a1b89d2b9a1101b54bcb2
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628008"
---
# <a name="troubleshoot-custom-functions"></a>Solução de problemas de funções personalizadas

Ao desenvolver funções personalizadas, você poderá encontrar erros no produto durante a criação e testes das funções.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Para resolver problemas, você pode [habilitar o log de tempo de execução para capturar erros](#enable-runtime-logging) e consultar as [mensagens de erro nativas do Excel](#check-for-excel-error-messages). Além disso, verifique se há erros comuns, como [deixar promessas não resolvidas](#ensure-promises-return) e esquecer de [associar as funções](#my-functions-wont-load-associate-functions).

## <a name="enable-runtime-logging"></a>Habilitar o log de tempo de execução

Se estiver testando o suplemento do Office no Windows, você deverá [habilitar o log de tempo de execução](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in). O log de tempo de execução entrega instruções `console.log` a um arquivo de log separado criado para ajudar você a descobrir problemas. As instruções abrangem vários erros, incluindo os relacionados ao arquivo de manifesto XML do suplemento, condições do tempo de execução ou a instalação de funções personalizadas.  Saiba mais sobre o log de tempo de execução em [Usar o log de tempo de execução para depurar seu suplemento](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).  

### <a name="check-for-excel-error-messages"></a>Verificar se há mensagens de erro do Excel

O Excel tem diversas mensagens de erro internas que serão retornadas para uma célula se houver um erro de cálculo. As funções personalizadas usam apenas as seguintes mensagens de erro: `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A` e `#BUSY!`.

Geralmente, estes erros correspondem aos erros que você já deve estar familiarizado no Excel. Existem apenas algumas exceções específicas para as funções personalizadas, listadas aqui:

- Um erro `#NAME` geralmente significa que houve um problema ao registrar as suas funções.
- Um erro `#VALUE` normalmente indica um erro no arquivo de script das funções.
- Um erro `#N/A` também pode ser um sinal de que esta função, embora registrada, não pode ser executada. Isto é normalmente devido à um comando `CustomFunctions.associate` em falta.
- Um erro `#REF!` pode indicar que o nome da sua função é o mesmo nome de uma função em um suplemento já existente.

## <a name="clear-the-office-cache"></a>Limpar o cache do Office

Informações sobre funções personalizadas são armazenadas em cache pelo Office. Às vezes, ao desenvolver e recarregar repetidamente um suplemento com funções personalizadas, as suas alterações podem não aparecer. Isso pode ser corrigido limpando o cache do Office. Para mais informações, consulte a seção «Limpar o Cache do Office» no artigo [Validar e solucionar problemas com seu manifesto](https://docs.microsoft.com/office/dev/add-ins/testing/troubleshoot-manifest?branch=master#clear-the-office-cache)

## <a name="common-issues"></a>Problemas comuns

### <a name="my-functions-wont-load-associate-functions"></a>Minhas funções não carregam: associar funções

No arquivo de script das funções personalizadas, você precisa associar cada função personalizada à respectiva ID especificada no [arquivo de metadados JSON](custom-functions-json.md). Isso é feito usando o método `CustomFunctions.associate()`. Normalmente, essa chamada de método é feita após cada função ou no final do arquivo de script. Se uma função personalizada não estiver associada, ele não funcionará.

O exemplo a seguir mostra uma função add, seguida pelo nome `add` da função que está sendo associada a `ADD` da id JSON correspondente.

```js
/**
 * Add two numbers.
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

Saiba mais sobre esse processo em [Associar os nomes de função com metadados JSON](/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata).

### <a name="cant-open-add-in-from-localhost-use-a-local-loopback-exception"></a>Não é possível abrir um suplemento de um localhost: utilize uma exceção de loopback local

Se você vir o erro "Não é possível abrir este suplemento de um localhost", será necessário habilitar uma exceção de loopback local. Para obter detalhes sobre como fazer isso, confira [este artigo de suporte da Microsoft](https://support.microsoft.com/pt-BR/help/4490419/local-loopback-exemption-does-not-work).

### <a name="ensure-promises-return"></a>Garantir que as promessas retornem resultados

Quando o Excel está aguardando a conclusão de uma função personalizada, ele exibe #BUSY! na célula. Se o código da função personalizada retornar uma promessa, mas a promessa não retornar um resultado, o Excel continuará exibindo #BUSY!. Verifique suas funções para garantir que as promessas estejam retornando corretamente um resultado para uma célula.

### <a name="error-the-dev-server-is-already-running-on-port-3000"></a>Erro: O servidor de desenvolvimento já está em execução na porta 3000

Às vezes, ao executar `npm start` você poderá ver um erro que o servidor de desenvolvimento já está executando na porta 3000 (ou qualquer outra porta que o seu suplemento use). Você pode parar o servidor de desenvolvimento executando `npm stop` ou fechando a janela Node.js. Mas em alguns casos, poderá levar alguns minutos para que o servidor de desenvolvimento realmente pare de executar.

## <a name="reporting-feedback"></a>Fornecer comentários

Se você tiver problemas que não estão descritos aqui, fale conosco. Há duas maneiras de relatar problemas.

### <a name="in-excel-on-windows-or-mac"></a>No Excel para Windows ou Mac

Se estiver usando o Excel para Windows ou Mac, envie comentários à equipe de extensibilidade do Office diretamente do Excel. Para fazer isso, selecione **Arquivo -> Comentários -> Enviar um Rosto Triste**. Enviando um Rosto Triste, você fornece os registros necessários para entendermos o problema que você está enfrentando.

### <a name="in-github"></a>No Github

Sinta-se à vontade para enviar problemas encontrados através do recurso "Comentários do conteúdo" na parte inferior de todas as páginas de documentação ou [informe um novo problema diretamente no repositório de funções personalizadas](https://github.com/OfficeDev/Excel-Custom-Functions/issues).

## <a name="next-steps"></a>Próximas etapas
Saiba como [depurar as suas funções personalizadas](custom-functions-debugging.md).

## <a name="see-also"></a>Confira também

* [Geração automática de metadados das funções personalizadas](custom-functions-json-autogeneration.md)
* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
* [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).
* [Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário](make-custom-functions-compatible-with-xll-udf.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
