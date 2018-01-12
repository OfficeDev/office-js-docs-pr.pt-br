# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a>Validar e solucionar problemas com seu manifesto

Use esses métodos para validar e solucionar problemas em seu manifesto. 

- [Validar o manifesto de Suplementos do Office com o Validador de Suplemento do Office](validate-the-office-add-ins-manifest-against-validator)   
- [Validar o manifesto de Suplementos do Office em relação ao esquema XML](validate-the-office-add-ins-manifest-against-the-xml-schema)
- [Use o log de tempo de execução para depurar o manifesto do suplemento do Office](use-runtime-logging-to-debug-the-manifest-for-your-office-add-in)

## <a name="validate-your-manifest-with-the-office-add-in-validator"></a>Validar o manifesto com o Validador de Suplemento do Office
Para ajudar a garantir que o arquivo de manifesto que descreve o Suplemento do Office está correto e completo, valide-o com base no [Validador de Suplemento do Office](https://github.com/OfficeDev/office-addin-validator).

Para usar o Validador de Suplemento do Office para validar o manifesto:

1. Instale o [Node.js](https://nodejs.org/download/). 
2. Abra um prompt de comando/terminal como administrador e instale o Validador de Suplemento do Office e as respectivas dependências globalmente usando o seguinte comando:

    ```
    npm install -g office-addin-validator
    ```
    
    > **Observação:** se já instalou o Office, atualize para a versão mais recente para que o validador seja instalado como uma dependência.

3. Para validar o manifesto, execute o seguinte comando: Substitua MANIFEST.XML pelo caminho para o arquivo XML de manifesto.

    ```
    validate-office-addin MANIFEST.XML
    ```


## <a name="validate-your-manifest-against-the-xml-schema"></a>Validar seu manifesto em relação ao esquema XML

Para ajudar a garantir que o arquivo de manifesto segue o esquema correto, valide-o com base nos arquivos [XSD (definição de esquema XML)](https://github.com/OfficeDev/office-js-docs/tree/master/docs/overview/schemas). Você pode usar uma ferramenta de validação para executar essa validação. 

Para usar uma ferramenta de validação de esquema XML da linha de comando para validar seu manifesto:

1.  Instale o [tar](https://www.gnu.org/software/tar/) e o [libxml](http://xmlsoft.org/FAQ.html), caso ainda não tenha o feito. 
2.  Execute o seguinte comando. Substitua XSD_FILE pelo caminho para o arquivo XSD do manifesto e XML_FILE pelo caminho para o arquivo XML do manifesto.
    ```
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="use-runtime-logging-to-debug-your-add-in-manifest"></a>Usar o log do tempo de execução para depurar seu manifesto de suplemento

Você pode usar o log do tempo de execução para depurar o manifesto do seu suplemento. Esse recurso pode ajudá-lo a identificar e corrigir problemas com seu manifesto que não são detectados pela validação de esquema XSD, como uma incompatibilidade entre as identificações dos recursos. O log do tempo de execução é particularmente útil para depurar suplementos que implementam comandos de suplemento.  

>**Observação:** o recurso de log de tempo de execução está atualmente disponível para o Office 2016 para área de trabalho.

### <a name="turn-on-runtime-logging"></a>Ativar o log de tempo de execução

>**Importante:** o log do tempo de execução afeta o desempenho. Ative-o somente quando precisar depurar problemas com seu manifesto de suplemento.

1. Verifique se você está executando o Office 2016 para área de trabalho na compilação **16.0.7019** ou posterior. 
2. Adicione a chave de registro `RuntimeLogging` sob 'HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\'. 
3. Defina o valor padrão da chave para o caminho completo do arquivo onde você deseja que o log seja gravado. Para obter um exemplo, veja [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip). 

 > **Observação:** A pasta na qual o arquivo de log será gravado deverá existir e você deverá ter permissões de gravação. 
 
A imagem a seguir mostra qual deve ser a aparência de registro. ![Captura de tela do editor do registro com uma chave de registro RuntimeLogging](http://i.imgur.com/Sa9TyI6.png)

Para desativar o recurso, remova a chave do registro `RuntimeLogging`. 

### <a name="troubleshoot-issues-with-your-manifest"></a>Solucionar problemas com o manifesto

Para usar o log do tempo de execução para solucionar problemas ao carregar um suplemento:
 
1. [Realize o sideload do seu suplemento](sideload-office-add-ins-for-testing.md) para teste. 

    >Observação: recomenda-se fazer o sideload apenas do suplemento que você está testando para minimizar a quantidade de mensagens no arquivo de log.
2. Se nada acontecer e você não vir seu suplemento (e ele não estiver aparecendo na caixa de diálogo de suplementos), abra o arquivo de log.
3. Procure pela ID de seu suplemento no arquivo de log, definida no seu manifesto. No arquivo de log, essa ID está marcada como `SolutionId`. 

No exemplo a seguir, o arquivo de log identifica um controle que aponta para um arquivo de recurso que não existe. Neste exemplo, a correção seria reparar o erro de digitação no manifesto ou adicionar o recurso que está faltando.

![Captura de tela de um arquivo de log com uma entrada que especifica uma identificação de recurso que não foi encontrado](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a>Problemas conhecidos com o log de tempo de execução

Talvez você veja mensagens no arquivo de log que são confusas ou que estão classificadas incorretamente. Por exemplo:

- A mensagem `Medium   Current host not in add-in's host list` seguida por `Unexpected Parsed manifest targeting different host` é incorretamente classificada como um erro.
- Se você vir a mensagem `Unexpected    Add-in is missing required manifest fields  DisplayName` e ela não contiver uma SolutionId, o erro provavelmente não está relacionado ao suplemento que você está depurando. 
- Todas as mensagens `Monitorable` indicam erros esperados do ponto de vista do sistema. Às vezes, indicam um problema com o seu manifesto, como um elemento que foi soletrado incorretamente e que foi ignorado, mas que não fez com que o manifesto falhasse. 

## <a name="clear-the-office-cache"></a>Limpar o cache do Office

Se alterações que você fez no manifesto, como nomes de arquivo dos ícones de botão da faixa de opções ou o texto de comandos de suplemento, não parecerem entrar em vigor, tente limpar o cache do Office no computador. 

#### <a name="for-windows"></a>No Windows:
Exclua o conteúdo da pasta `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.

#### <a name="for-mac"></a>No Mac:
Exclua o conteúdo da pasta `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.

#### <a name="for-ios"></a>No iOS:
Chame `window.location.reload(true)` usando o JavaScript no suplemento para forçar um recarregamento. Outra alternativa é reinstalar o Office.

## <a name="additional-resources"></a>Recursos adicionais

- [Manifesto XML dos Suplementos do Office](../overview/add-in-manifests.md)
- [Realizar sideload de suplementos do Office para teste](sideload-office-add-ins-for-testing.md)
- [Depurar suplementos do Office](debug-add-ins-using-f12-developer-tools-on-windows-10.md)