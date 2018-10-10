# <a name="method-element"></a>Elemento Method

Especifica um método individual a partir da API do JavaScript para Office que o Suplemento do Office exige para ativar.

**Tipo de suplemento:** Conteúdo, Painel de tarefas

## <a name="syntax"></a>Sintaxe

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>Contido em

[Métodos](methods.md)

## <a name="attributes"></a>Atributos

|**Atributo**|**Tipo**|**Obrigatório**|**Descrição**|
|:-----|:-----|:-----|:-----|
|Nome|cadeia de caracteres|obrigatório|Especifica o nome do método necessário qualificado com seu objeto pai. Por exemplo, para especificar o método **getSelectedDataAsync**, você deve especificar `"Document.getSelectedDataAsync"`.|

## <a name="remarks"></a>Comentários

Os elementos **Methods** e **Method** não são compatíveis com os suplementos de email. Para obter mais informações sobre os conjuntos de requisitos, consulte [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

> [!IMPORTANT] 
> Como não há forma de especificar o requisito de versão mínima de métodos individuais, para verificar se um método está disponível em tempo de execução, você também deve usar uma instrução **if** ao chamar o método no script do suplemento. Para obter mais informações sobre como fazer isso, consulte [Noções básicas sobre a API do JavaScript para Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).

