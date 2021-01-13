---
ms.date: 01/08/2020
description: Solucionar problemas comuns com funções personalizadas do Excel.
title: Solução de problemas das funções personalizadas
localization_priority: Normal
ms.openlocfilehash: d9f912b1cd98b04c6d0e207c79491313dc794719
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839835"
---
# <a name="troubleshoot-custom-functions"></a>Solução de problemas de funções personalizadas

Ao desenvolver funções personalizadas, você poderá encontrar erros no produto durante a criação e testes das funções.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Para resolver problemas, você pode [habilitar o log de tempo de execução para capturar erros](#enable-runtime-logging) e consultar as [mensagens de erro nativas do Excel](#check-for-excel-error-messages). Alem disso, verifique se há erros comuns, como [deixar promessas não resolvidas](#ensure-promises-return).

## <a name="enable-runtime-logging"></a>Habilitar o log de tempo de execução

Se estiver testando o suplemento do Office no Windows, você deverá [habilitar o log do tempo de execução](../testing/runtime-logging.md). O log de tempo de execução entrega instruções `console.log` a um arquivo de log separado criado para ajudar você a descobrir problemas. As instruções abrangem vários erros, incluindo os relacionados ao arquivo de manifesto XML do suplemento, condições do tempo de execução ou a instalação de funções personalizadas. Para saber mais sobre o log do tempo de execução, confira [Depurar seu suplemento com o log do tempo de execução](../testing/runtime-logging.md).

### <a name="check-for-excel-error-messages"></a>Verificar se há mensagens de erro do Excel

O Excel tem diversas mensagens de erro internas que serão retornadas para uma célula se houver um erro de cálculo. As funções personalizadas usam apenas as seguintes mensagens de erro: `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A` e `#BUSY!`.

Geralmente, estes erros correspondem aos erros que você já deve estar familiarizado no Excel. Existem apenas algumas exceções específicas para as funções personalizadas, listadas aqui:

- Um erro `#NAME` geralmente significa que houve um problema ao registrar as suas funções.
- Um erro `#N/A` também pode ser um sinal de que esta função, embora registrada, não pode ser executada. Isto é normalmente devido à um comando `CustomFunctions.associate` em falta.
- Um `#VALUE` erro normalmente indica um erro no arquivo de script das funções.
- Um erro `#REF!` pode indicar que o nome da sua função é o mesmo nome de uma função em um suplemento já existente.

## <a name="clear-the-office-cache"></a>Limpar o cache do Office

Informações sobre funções personalizadas são armazenadas em cache pelo Office. Às vezes, ao desenvolver e recarregar repetidamente um suplemento com funções personalizadas, as suas alterações podem não aparecer. Isso pode ser corrigido limpando o cache do Office. Para saber mais, confira [Limpar o cache do Office](../testing/clear-cache.md).

## <a name="common-problems-and-solutions"></a>Problemas comuns e soluções

### <a name="cant-open-add-in-from-localhost-use-a-local-loopback-exception"></a>Não é possível abrir um suplemento de um localhost: utilize uma exceção de loopback local

Se você vir o erro "Não é possível abrir este suplemento de um localhost", será necessário habilitar uma exceção de loopback local. Para obter detalhes sobre como fazer isso, confira [este artigo de suporte da Microsoft](https://support.microsoft.com/help/4490419/local-loopback-exemption-does-not-work).

### <a name="runtime-logging-reports-typeerror-network-request-failed-on-excel-on-windows"></a>Relatórios de log de tempo de execução "TypeError: Falha na solicitação de rede" no Excel para Windows

Se você ver o erro "TypeError: Falha na solicitação de rede" em seu [log de tempo de execução](custom-functions-troubleshooting.md#enable-runtime-logging) enquanto faz chamadas para seu servidor localhost, você precisará habilitar uma exceção de loopback local. Para mais detalhes sobre como fazer isso, confira *Opção #2* neste [artigo de suporte da Microsoft](https://support.microsoft.com/help/4490419/local-loopback-exemption-does-not-work).

### <a name="ensure-promises-return"></a>Garantir que as promessas retornem resultados

Quando o Excel está aguardando a conclusão de uma função personalizada, ele exibe #BUSY! na célula. Se o código da função personalizada retornar uma promessa, mas a promessa não retornar um resultado, o Excel continuará exibindo `#BUSY!`. Verifique suas funções para garantir que as promessas estejam retornando corretamente um resultado para uma célula.

### <a name="error-the-dev-server-is-already-running-on-port-3000"></a>Erro: O servidor de desenvolvimento já está em execução na porta 3000

Às vezes, ao executar `npm start` você poderá ver um erro que o servidor de desenvolvimento já está executando na porta 3000 (ou qualquer outra porta que o seu suplemento use). Você pode parar o servidor de desenvolvimento executando `npm stop` ou fechando a janela Node.js. Em alguns casos, pode levar alguns minutos para que o servidor dev pare de ser executado.

### <a name="my-functions-wont-load-associate-functions"></a>Minhas funções não carregam: associar funções

Nos casos em que seu JSON não tiver sido registrado e você tiver criado os seus próprios metadados JSON, talvez receba um `#VALUE!`erro ou receba uma notificação de que o seu suplemento não pode ser carregado. Geralmente, isso significa que você precisa associar cada função personalizada a `id`propriedade especificada no [arquivo de metadados JSON](custom-functions-json.md). Isso é feito usando o método `CustomFunctions.associate()`. Normalmente, essa chamada de método é feita após cada função ou no final do arquivo de script. Se uma função personalizada não estiver associada, ele não funcionará.

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

Para obter mais informações sobre esse processo, consulte [Associando nomes de função com metadados JSON.](../excel/custom-functions-json.md#associating-function-names-with-json-metadata)

## <a name="known-issues"></a>Problemas conhecidos

Problemas conhecidos são rastreados e relatados no repositório GitHub de funções [personalizadas do Excel.](https://github.com/OfficeDev/Excel-Custom-Functions/issues)

## <a name="reporting-feedback"></a>Fornecer comentários

Se você tiver problemas que não estão descritos aqui, fale conosco. Há duas maneiras de relatar problemas.

### <a name="in-excel-on-windows-or-mac"></a>No Excel para Windows ou Mac

Se estiver usando o Excel no Windows ou Mac, envie comentários à equipe de extensibilidade do Office diretamente do Excel. Para fazer isso, selecione **Arquivo -> Comentários -> Enviar um Rosto Triste**. Enviando um Rosto Triste, você fornece os registros necessários para entendermos o problema que você está enfrentando.

### <a name="in-github"></a>No Github

Sinta-se à vontade para enviar problemas encontrados através do recurso "Comentários do conteúdo" na parte inferior de todas as páginas de documentação ou [informe um novo problema diretamente no repositório de funções personalizadas](https://github.com/OfficeDev/Excel-Custom-Functions/issues).

## <a name="next-steps"></a>Próximas etapas
Saiba como [tornar as suas funções personalizadas compatíveis com as funções definidas pelo usuário de XLL](make-custom-functions-compatible-with-xll-udf.md).

## <a name="see-also"></a>Confira também

* [Gerar metadados JSON automaticamente para funções personalizadas](custom-functions-json-autogeneration.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
