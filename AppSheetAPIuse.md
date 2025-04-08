# Guide d'utilisation de l'API AppSheet

## Introduction

AppSheet est une plateforme de développement d'applications sans code qui permet de créer des applications mobiles et web à partir de sources de données comme Google Sheets, Excel, ou des bases de données. L'API AppSheet v2 permet d'interagir programmatiquement avec les applications AppSheet, offrant ainsi la possibilité d'intégrer les fonctionnalités d'AppSheet dans d'autres systèmes ou applications.

Ce document détaille l'utilisation de l'API AppSheet v2 telle qu'implémentée dans notre serveur MCP (Model Context Protocol), qui sert d'interface entre les modèles de langage (LLMs) et les applications AppSheet.

## Configuration et authentification

### Prérequis

Pour utiliser l'API AppSheet, vous avez besoin de :

1. Une application AppSheet existante
2. L'API activée pour cette application (dans les paramètres de l'éditeur AppSheet)
3. Un identifiant d'application (App ID)
4. Une clé d'accès API (Access Key)

### Configuration

Les informations d'authentification sont généralement stockées dans des variables d'environnement :

```
APPSHEET_APP_ID=votre_app_id
APPSHEET_ACCESS_KEY=votre_cle_acces
```

### Authentification

L'authentification à l'API AppSheet se fait via un en-tête HTTP `ApplicationAccessKey` qui contient la clé d'accès API :

```javascript
const response = await axios.post(url, requestBody, {
  headers: {
    'ApplicationAccessKey': APPSHEET_ACCESS_KEY,
    'Content-Type': 'application/json'
  }
});
```

## Structure de l'API

### URL de base

L'URL de base pour l'API AppSheet v2 est :

```
https://api.appsheet.com/api/v2
```

### Points d'accès principaux

Les opérations sur les tables se font via le point d'accès :

```
/apps/{appId}/tables/{tableName}/Action
```

Où :
- `{appId}` est l'identifiant de votre application AppSheet
- `{tableName}` est le nom de la table sur laquelle vous souhaitez effectuer une opération
- `Action` indique que vous allez effectuer une action sur cette table

## Opérations CRUD

L'API AppSheet permet d'effectuer les opérations CRUD (Create, Read, Update, Delete) standard sur les données de vos applications.

### Recherche d'enregistrements (Read)

Pour rechercher des enregistrements dans une table, utilisez l'action `Find` :

```javascript
// Structure de la requête
const requestBody = {
  Action: "Find",
  Properties: {
    Filter: "nom = 'Dupont'",  // Optionnel
    PageSize: 10,              // Optionnel
    PageToken: "token123"      // Optionnel
  }
};
```

#### Formats de réponse

L'API AppSheet peut renvoyer les résultats dans deux formats différents :

1. **Format tableau direct** : Un tableau d'objets représentant les enregistrements

```json
[
  { "id": "1", "name": "Record 1", "value": 100 },
  { "id": "2", "name": "Record 2", "value": 200 }
]
```

2. **Format avec Rows et Properties** : Un objet contenant un tableau `Rows` et un objet `Properties`

```json
{
  "Rows": [
    { "id": "1", "name": "Record 1", "value": 100 },
    { "id": "2", "name": "Record 2", "value": 200 }
  ],
  "Properties": {
    "NextPageToken": "token123"
  }
}
```

Le format avec `Rows` et `Properties` est généralement utilisé lorsque des informations supplémentaires (comme un jeton de pagination) doivent être incluses dans la réponse.

#### Filtrage

Le filtrage se fait via la propriété `Filter` dans l'objet `Properties` de la requête. La syntaxe de filtrage est similaire à celle utilisée dans les formules AppSheet :

```javascript
// Exemples de filtres
"nom = 'Dupont'"                  // Égalité
"age > 30"                        // Comparaison numérique
"ville IN ('Paris', 'Lyon')"      // Liste de valeurs
"nom LIKE 'Du%'"                  // Recherche par motif
"dateCreation > TODAY()"          // Fonctions de date
"statut = 'Actif' AND age > 18"   // Opérateurs logiques
```

#### Pagination

La pagination se fait via les propriétés `PageSize` et `PageToken` :

- `PageSize` : Nombre maximum d'enregistrements à retourner
- `PageToken` : Jeton de pagination pour récupérer la page suivante

La réponse peut inclure un jeton `NextPageToken` dans l'objet `Properties` pour récupérer la page suivante :

```javascript
// Requête pour la première page
const requestBody = {
  Action: "Find",
  Properties: {
    PageSize: 10
  }
};

// Réponse
{
  "Rows": [...],
  "Properties": {
    "NextPageToken": "token123"
  }
}

// Requête pour la page suivante
const requestBody = {
  Action: "Find",
  Properties: {
    PageSize: 10,
    PageToken: "token123"
  }
};
```

### Ajout d'enregistrements (Create)

Pour ajouter un nouvel enregistrement, utilisez l'action `Add` :

```javascript
// Structure de la requête
const requestBody = {
  Action: "Add",
  Rows: [
    {
      nom: "Dupont",
      prenom: "Jean",
      email: "jean.dupont@example.com"
    }
  ]
};
```

#### Format de réponse

La réponse contient généralement l'enregistrement ajouté, potentiellement avec des valeurs calculées ou générées automatiquement (comme un ID) :

```json
{
  "Rows": [
    {
      "id": "3",
      "nom": "Dupont",
      "prenom": "Jean",
      "email": "jean.dupont@example.com",
      "dateCreation": "2025-04-08T08:30:00Z"
    }
  ]
}
```

### Modification d'enregistrements (Update)

Pour modifier un enregistrement existant, utilisez l'action `Edit` :

```javascript
// Structure de la requête
const requestBody = {
  Action: "Edit",
  Rows: [
    {
      id: "3",  // Clé primaire pour identifier l'enregistrement
      email: "nouveau.email@example.com"
    }
  ]
};
```

Il est important d'inclure la clé primaire dans les données pour identifier l'enregistrement à modifier. Vous n'avez besoin d'inclure que les champs que vous souhaitez modifier, pas tous les champs de l'enregistrement.

#### Format de réponse

La réponse contient généralement l'enregistrement modifié :

```json
{
  "Rows": [
    {
      "id": "3",
      "nom": "Dupont",
      "prenom": "Jean",
      "email": "nouveau.email@example.com",
      "dateModification": "2025-04-08T08:35:00Z"
    }
  ]
}
```

### Suppression d'enregistrements (Delete)

Pour supprimer un enregistrement, utilisez l'action `Delete` :

```javascript
// Structure de la requête
const requestBody = {
  Action: "Delete",
  Rows: [
    {
      id: "3"  // Clé primaire pour identifier l'enregistrement
    }
  ]
};
```

Comme pour la modification, vous devez inclure la clé primaire pour identifier l'enregistrement à supprimer.

#### Format de réponse

La réponse pour une suppression réussie est généralement simple :

```json
{
  "Status": "Success"
}
```

## Actions personnalisées

AppSheet permet de définir des actions personnalisées dans vos applications. Ces actions peuvent être invoquées via l'API en utilisant l'action `Action` :

```javascript
// Structure de la requête
const requestBody = {
  Action: "Action",
  Rows: [
    { id: "1" },
    { id: "2" }
  ],
  Properties: {
    ActionName: "EnvoyerEmail",
    Format: "HTML",
    Cc: "support@example.com"
  }
};
```

Dans cet exemple :
- `Rows` contient les identifiants des enregistrements sur lesquels l'action doit être effectuée
- `Properties.ActionName` spécifie le nom de l'action à invoquer
- Les autres propriétés dans `Properties` sont des paramètres spécifiques à l'action

#### Format de réponse

La réponse pour une action réussie est généralement simple :

```json
{
  "Status": "Success"
}
```

Certaines actions peuvent renvoyer des données supplémentaires spécifiques à l'action.

## Gestion des erreurs

L'API AppSheet renvoie des erreurs HTTP standard avec des codes d'état appropriés :

- `400 Bad Request` : Requête mal formée ou paramètres invalides
- `401 Unauthorized` : Authentification invalide
- `403 Forbidden` : Authentification valide mais accès refusé (par exemple, API non activée)
- `404 Not Found` : Ressource non trouvée (par exemple, table inexistante)
- `500 Internal Server Error` : Erreur interne du serveur AppSheet

Les réponses d'erreur contiennent généralement des informations détaillées sur l'erreur :

```json
{
  "type": "https://tools.ietf.org/html/rfc9110#section-15.5.4",
  "title": "Forbidden",
  "status": 403,
  "detail": "REST API invoke request failed: The API is not enabled for the called application on the Editor's Settings > Integrations > In tab.",
  "traceId": "00-f2cb2229866d2d6b964c7a124ace9f1c-262168ec28f44b8c-00"
}
```

## Bonnes pratiques

### Validation des paramètres

Validez toujours les paramètres requis avant d'envoyer une requête à l'API AppSheet :

```javascript
if (!tableName) {
  throw new Error('Le paramètre tableName est requis');
}
```

### Gestion des formats de réponse

Comme mentionné précédemment, l'API AppSheet peut renvoyer les résultats dans différents formats. Assurez-vous de gérer ces différents formats dans votre code :

```javascript
let records = [];
if (Array.isArray(result)) {
  // Format tableau direct
  records = result;
} else if (result.Rows) {
  // Format avec Rows
  records = result.Rows;
}
```

### Pagination

Pour récupérer de grandes quantités de données, utilisez la pagination pour éviter de surcharger l'API et votre application :

```javascript
let allRecords = [];
let pageToken = null;

do {
  const result = await findRecords(tableName, filter, pageSize, pageToken);
  
  if (Array.isArray(result)) {
    allRecords = allRecords.concat(result);
    pageToken = null; // Pas de pagination dans ce format
  } else if (result.Rows) {
    allRecords = allRecords.concat(result.Rows);
    pageToken = result.Properties?.NextPageToken;
  }
} while (pageToken);
```

### Gestion des erreurs

Implémentez une gestion robuste des erreurs pour traiter les différentes erreurs qui peuvent survenir lors de l'utilisation de l'API AppSheet :

```javascript
try {
  const result = await makeRequest(tableName, action, data);
  // Traiter le résultat
} catch (error) {
  if (error.response?.status === 403) {
    // Gérer l'erreur d'accès refusé
    console.error("L'API n'est pas activée pour cette application");
  } else if (error.response?.status === 404) {
    // Gérer l'erreur de ressource non trouvée
    console.error(`La table '${tableName}' n'existe pas`);
  } else {
    // Gérer les autres erreurs
    console.error("Erreur lors de la requête API:", error.message);
  }
}
```

## Exemples d'utilisation

### Recherche d'enregistrements avec filtrage

```javascript
// Recherche de clients à Paris
const result = await appsheetClient.findRecords(
  "Clients",
  "ville = 'Paris'"
);

// Traitement des résultats
const clients = Array.isArray(result) ? result : result.Rows || [];
console.log(`${clients.length} clients trouvés à Paris`);
```

### Ajout d'un nouvel enregistrement

```javascript
// Ajout d'un nouveau client
const newClient = {
  nom: "Dupont",
  prenom: "Jean",
  email: "jean.dupont@example.com",
  ville: "Paris",
  dateInscription: new Date().toISOString()
};

const result = await appsheetClient.addRecord("Clients", newClient);
const clientId = result.Rows?.[0]?.id;
console.log(`Nouveau client ajouté avec l'ID ${clientId}`);
```

### Modification d'un enregistrement

```javascript
// Modification de l'email d'un client
const clientId = "C001";
const updates = {
  email: "nouveau.email@example.com"
};

await appsheetClient.editRecord("Clients", clientId, "id", updates);
console.log(`Email du client ${clientId} mis à jour`);
```

### Suppression d'un enregistrement

```javascript
// Suppression d'un client
const clientId = "C001";
await appsheetClient.deleteRecord("Clients", clientId, "id");
console.log(`Client ${clientId} supprimé`);
```

### Invocation d'une action personnalisée

```javascript
// Envoi d'une facture à plusieurs clients
const clientIds = ["C001", "C002", "C003"];
await appsheetClient.invokeAction(
  "Clients",
  "EnvoyerFacture",
  clientIds,
  "id",
  { format: "PDF", includeDetails: true }
);
console.log(`Factures envoyées à ${clientIds.length} clients`);
```

## Limitations et considérations

### Limites de débit

AppSheet peut imposer des limites de débit sur les appels API. Consultez la documentation officielle pour connaître les limites actuelles.

### Sécurité

- Ne stockez jamais votre clé d'accès API dans le code source ou dans des fichiers publics
- Utilisez des variables d'environnement ou des services de gestion de secrets pour stocker vos informations d'authentification
- Limitez les permissions de votre clé d'accès API au strict nécessaire

### Performance

- Utilisez le filtrage côté serveur plutôt que de récupérer toutes les données et de filtrer côté client
- Utilisez la pagination pour les grandes quantités de données
- Limitez les champs retournés si l'API le permet

## Ressources et références

- [Documentation officielle de l'API AppSheet](https://help.appsheet.com/en/articles/1979979-api-overview)
- [Centre d'aide AppSheet](https://help.appsheet.com/)
- [Communauté AppSheet](https://community.appsheet.com/)

## Dépannage

### Erreur "API not enabled"

Si vous recevez une erreur indiquant que l'API n'est pas activée, assurez-vous que l'API est activée dans les paramètres de votre application AppSheet :

1. Ouvrez l'éditeur AppSheet
2. Allez dans "Settings" > "Integrations" > "In"
3. Activez l'option "Enable API"

### Erreur "Invalid Access Key"

Si vous recevez une erreur d'authentification, vérifiez que :

1. Votre clé d'accès API est correcte
2. La clé n'a pas expiré
3. La clé a les permissions nécessaires pour effectuer l'opération demandée

### Erreur "Table not found"

Si vous recevez une erreur indiquant qu'une table n'existe pas, vérifiez que :

1. Le nom de la table est correctement orthographié
2. La table existe bien dans votre application AppSheet
3. Vous avez les permissions nécessaires pour accéder à cette table

## Conclusion

L'API AppSheet v2 offre une interface puissante pour interagir programmatiquement avec vos applications AppSheet. En suivant les bonnes pratiques et en comprenant la structure des requêtes et des réponses, vous pouvez intégrer efficacement les fonctionnalités d'AppSheet dans vos propres applications et systèmes.

Ce document a présenté les principales fonctionnalités de l'API AppSheet telles qu'implémentées dans notre serveur MCP, mais l'API peut offrir des fonctionnalités supplémentaires ou spécifiques à certaines applications. Consultez la documentation officielle pour des informations plus détaillées et à jour.
