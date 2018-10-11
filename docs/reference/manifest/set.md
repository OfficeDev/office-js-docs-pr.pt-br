# <a name="set-element"></a>Elemento Set

Especifica um conjunto de requisitos a partir da API JavaScript para Office que o seu suplemento do Office exige para ativar.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>Contido em

[Conjuntos](sets.md)

## <a name="attributes"></a>Atributos

|**Atributo**|**Tipo**|**Obrigatório**|**Descrição**|
|:-----|:-----|:-----|:-----|
|Nome|sequência de caracteres|obrigatório|O nome de um [conjunto de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).|
|MinVersion|sequência de caracteres|opcional|Especifica a versão mínima do conjunto de APIs exigido pelo seu suplemento. Substitui o valor de **DefaultMinVersion**, se ele estiver especificado no elemento [Sets](sets.md) pai.|

## <a name="remarks"></a>Comentários

Para obter mais informações sobre os conjuntos de requisitos, consulte [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Para saber mais sobre o atributo **MinVersion** do elemento **Set** e o atributo **DefaultMinVersion** do elemento **Sets**, confira [Definir o elemento Requirements no manifesto](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).

> [!IMPORTANT] 
> Para suplementos de email, há somente um `"Mailbox"` conjunto de requisitos disponível. Para suplementos de email, há somente um conjunto de requisitos  contenteditable="false" class="locked monad selfClosingTag">`"Mailbox"` disponível. Além disso, você não pode declarar suporte para métodos específicos em suplementos de email.
