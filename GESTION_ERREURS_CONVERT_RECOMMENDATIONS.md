# Gestion des Erreurs - Recommandations pour convert.js

**Date:** 14 juin 2026  
**Contexte:** Conversion GEDCOM vers JSON pour données généalogiques

---

## 1. État Actuel de convert.js

### 1.1 Situation Actuelle
- **Fichier:** `converterGedJson/convert.js`
- **Statut:** ✅ Fonctionne pour les données actuelles (Hrysio1.json)
- **Gestion d'erreurs:** ❌ Minimale/Absente
- **Points faibles:**
  1. Pas de try/catch autour du traitement des objets
  2. Pas de validation des données
  3. Pas de logs détaillés en cas de problème
  4. Crash silencieux possible sur données malformées

### 1.2 Structure Actuelle
```
Entrée: Fichier GEDCOM (.ged)
   ↓
FileLines (streaming par ligne)
   ↓
Ged2Json (transformation en JSON)
   ↓
parseKnownObjects() (organisation en graphe)
   ↓
Output: JSON-LD (affichage console)
```

---

## 2. Problèmes Identifiés

### 2.1 Pas de Gestion d'Erreurs au Niveau des Objets
**Localisation:** Fonction `parseKnownObjects()` (ligne ~330)

**Problème:**
```javascript
for (var i = 0; i < keys.length; i++) {
    var objId = keys[i];
    var obj = knownObjects[keys[i]];  // ← Peut être undefined
    if (obj['@type'] == 'foaf:Person') {
        // Traitement sans vérification
        var relationships = asArray(obj, 'bio:relationship');
        // ← Si bio:relationship mal formé → crash
    }
}
```

**Impact:** Données mal formées causent crash silencieux du script

### 2.2 Pas de Validation des Références
**Problème:**
```javascript
var fam = knownObjects[relationships[j]['@id']];  // ← Retourne undefined si pas trouvé
var tmp = findSpouses(fam, obj['@id']);  // ← findSpouses() crashe si fam est undefined
```

**Impact:** Références manquantes ou mal formées causent crash

### 2.3 Pas d'Information sur les Pièces Jointes (OBJE)
**Problème:** Aucun log ni tracking spécifique des OBJE

**Impact:** Difficile de diagnostiquer pourquoi certaines pièces jointes ne sont pas traitées

### 2.4 Pas de Sortie d'Erreurs Structurée
**Problème:** Seule output est le JSON ou rien (crash silencieux)

**Impact:** Impossible de savoir quelles données ont des problèmes

---

## 3. Solutions Recommandées

### 3.1 Ajouter Gestion d'Erreurs Globale

**Implémentation:**
```javascript
// Au début du fichier
var errorLog = {
    count: 0,
    errors: [],
    warnings: [],
    skipped: []
};

// Wrapper pour traitement sécurisé
function safeProcessObject(obj, callback) {
    try {
        if (!obj || typeof obj !== 'object') {
            errorLog.warnings.push({
                type: 'INVALID_OBJECT',
                message: 'Object is not valid',
                objectId: obj ? obj['@id'] : 'unknown'
            });
            return null;
        }
        return callback(obj);
    } catch (error) {
        errorLog.count++;
        errorLog.errors.push({
            objectId: obj['@id'] || 'unknown',
            objectType: obj['@type'] || 'unknown',
            errorType: error.name,
            message: error.message,
            stack: error.stack,
            timestamp: new Date().toISOString()
        });
        return null;
    }
}
```

### 3.2 Valider les Références Avant Utilisation

**Implémentation:**
```javascript
function findSpouses(fam, knownId) {
    // Ajouter validation
    if (!fam) {
        errorLog.warnings.push({
            type: 'MISSING_FAMILY',
            familyId: 'unknown',
            message: 'Family object is null or undefined'
        });
        return false;
    }
    
    if (typeof fam !== 'object') {
        errorLog.warnings.push({
            type: 'INVALID_FAMILY',
            familyId: String(fam),
            message: 'Family is not an object'
        });
        return false;
    }
    
    // ... reste du code
}
```

### 3.3 Ajouter Tracking Spécifique des OBJE

**Implémentation:**
```javascript
var objeStats = {
    totalFound: 0,
    byType: {},  // Compter par type d'élément
    perPerson: [],  // Lister les OBJE par personne
    missing: []  // Références OBJE manquantes
};

// Dans parseKnownObjects()
if (obj['OBJE']) {
    var objes = Array.isArray(obj['OBJE']) ? obj['OBJE'] : [obj['OBJE']];
    objeStats.totalFound += objes.length;
    
    objeStats.perPerson.push({
        id: obj['@id'],
        name: _display_name(obj),
        count: objes.length,
        objes: objes.map(o => ({
            id: o['@id'],
            prim: o['_PRIM']
        }))
    });
    
    // Vérifier que chaque OBJE existe
    objes.forEach(o => {
        if (!knownObjects[o['@id']]) {
            objeStats.missing.push({
                objId: obj['@id'],
                objeId: o['@id'],
                message: 'Referenced OBJE not found'
            });
        }
    });
}
```

### 3.4 Sortie d'Erreurs Structurée

**Implémentation:**
```javascript
fs.on('end', function() {
    parseKnownObjects();
    
    // Afficher le JSON principal
    console.log(JSON.stringify(graph, null, 2));
    
    // Afficher les statistiques et erreurs sur stderr
    if (errorLog.count > 0 || errorLog.warnings.length > 0) {
        console.error('\n========== ERREURS ET AVERTISSEMENTS ==========');
        console.error(`Nombre d'erreurs: ${errorLog.count}`);
        console.error(`Nombre d'avertissements: ${errorLog.warnings.length}`);
        
        if (errorLog.errors.length > 0) {
            console.error('\nERREURS:');
            errorLog.errors.forEach((err, idx) => {
                console.error(`  [${idx + 1}] ${err.objectId} (${err.objectType})`);
                console.error(`      Message: ${err.message}`);
            });
        }
        
        if (errorLog.warnings.length > 0) {
            console.error('\nAVERTISSEMENTS:');
            errorLog.warnings.forEach((warn, idx) => {
                console.error(`  [${idx + 1}] ${warn.type}: ${warn.message}`);
            });
        }
        
        console.error('\nSTATISTIQUES OBJE:');
        console.error(`  Total trouvé: ${objeStats.totalFound}`);
        console.error(`  Références manquantes: ${objeStats.missing.length}`);
        console.error(`  Personnes avec OBJE: ${objeStats.perPerson.length}`);
    }
});
```

---

## 4. Plan d'Implémentation

### Phase 1: Validation et Logging (Faible Risque)
- [ ] Ajouter validation d'objets
- [ ] Ajouter logs d'erreurs sans changer la logique
- [ ] Tester avec fichiers existants

**Effort:** 2-3 heures  
**Risque:** Minimal

### Phase 2: Gestion d'Erreurs (Moyen Risque)
- [ ] Ajouter try/catch dans parseKnownObjects()
- [ ] Ajouter try/catch dans findSpouses() et findParents()
- [ ] Tester avec fichiers avec données malformées

**Effort:** 4-5 heures  
**Risque:** Moyen (changements logiques)

### Phase 3: Tracking OBJE (Haute Priorité pour le Projet)
- [ ] Implémenter objeStats
- [ ] Comparer avec fichier source GEDCOM original
- [ ] Documenter écarts entre source et output

**Effort:** 3-4 heures  
**Risque:** Minimal (ajout de tracking)

---

## 5. Exemple de Sortie Améliorée

### Avant (Actuellement)
```
[Affichage du JSON seul, ou crash silencieux]
```

### Après (Avec Améliorations)
```json
{
  "@context": {...},
  "@graph": [...]
}

========== ERREURS ET AVERTISSEMENTS ==========
Nombre d'erreurs: 0
Nombre d'avertissements: 2

AVERTISSEMENTS:
  [1] MISSING_FAMILY: Family _:8000 not found for person _:737
  [2] INVALID_REFERENCE: Referenced person _:9999 not found in index

STATISTIQUES OBJE:
  Total trouvé: 22
  Références manquantes: 0
  Personnes avec OBJE: 4
  
  • _:731 (Didier Michel BEDER): 0 OBJE
  • _:732 (Jeannine HRYSIO): 1 OBJE
  • _:737 (Alfred Édouard BEDER): 2 OBJE [PRIM: 1]
  • _:738 (Marie BOEGLER): 0 OBJE
```

---

## 6. Comparaison: convert.js vs generator.py

| Aspect | convert.js | generator.py |
|--------|-----------|-------------|
| **Gestion d'erreurs** | ❌ Aucune | ✅ try/except global |
| **Logging** | ❌ Aucun | ✅ Messages détaillés |
| **Validation** | ❌ Aucune | ✅ Validation des types |
| **Récupération** | ❌ Crash | ✅ Continue on error |
| **Tracking OBJE** | ❌ Aucun | ✅ Affichage OBJE |

**Recommandation:** Harmoniser convert.js avec les patterns de generator.py

---

## 7. Test de Validation

### Test 1: Données Valides (Regression)
```bash
node convert.js Hrysio-Export.ged > output.json
# Vérifier: même JSON qu'avant, 0 erreurs
```

### Test 2: Données Manquantes
```bash
# Modifier GEDCOM pour supprimer une famille référencée
# Vérifier: 1 avertissement MISSING_FAMILY
```

### Test 3: Données Malformées
```bash
# Ajouter données invalides au GEDCOM
# Vérifier: Erreur loggée, pas de crash
```

### Test 4: Analyse OBJE
```bash
# Compter OBJE en entrée vs sortie
# Vérifier: Nombres correspondent
```

---

## 8. Documentation Recommandée

Ajouter un fichier `CONVERT_ERROR_CODES.md`:
```markdown
# Codes d'Erreur convert.js

## ERRORS

### INVALID_OBJECT
- **Description:** Un objet n'est pas valide (null, undefined, ou pas un objet)
- **Action:** Vérifier le fichier GEDCOM source
- **Exemple:** Reference to null or invalid object

### MISSING_FAMILY
- **Description:** Une famille référencée par une personne n'existe pas
- **Action:** Vérifier la cohérence du fichier GEDCOM
- **Exemple:** Family _:999 referenced but not found

## WARNINGS

### MISSING_REFERENCE
- **Description:** Une référence croisée manque dans l'index
- **Action:** Peut être normal pour certains fichiers GEDCOM
- **Exemple:** Person _:123 has no spouse data
```

---

## 9. Checklist d'Implémentation

- [ ] Créer branche feature/error-handling
- [ ] Implémenter Phase 1 (Validation & Logging)
- [ ] Tester avec Hrysio1.json (regression)
- [ ] Implémenter Phase 2 (Error Handling)
- [ ] Tester avec données problématiques
- [ ] Implémenter Phase 3 (OBJE Tracking)
- [ ] Comparer statistiques avec source GEDCOM
- [ ] Rédiger documentation des codes d'erreur
- [ ] Créer pull request avec tests
- [ ] Merger dans main avec tous les tests passing

---

## 10. Priorités de Mise en Œuvre

**🔴 URGENT:** Phase 3 - Tracking OBJE
- Raison: Tu mentionnes 70 pièces jointes attendues vs 2 actuelles
- Action: Diagnostiquer où sont les 68 pièces jointes manquantes

**🟡 IMPORTANT:** Phase 1 - Validation & Logging
- Raison: Facilite diagnostique des problèmes futurs
- Action: À faire avant Phase 2

**🟢 SOUHAITABLE:** Phase 2 - Gestion d'erreurs complète
- Raison: Rend robuste et fiable
- Action: À faire quand Phase 1 et 3 sont complétées

---

## 11. Contact et Maintenance

Fichiers concernés:
- `converterGedJson/convert.js` - Main conversion script
- `converterGedJson/ged2Json.js` - JSON stream parser
- `python_poc/generator.py` - Post-processing (déjà amélioré ✅)

Documentation liée:
- `MODIFICATIONS_DOCUMENTATION.md` - Changements apportés
- `CONVERT_ERROR_CODES.md` - À créer avec codes d'erreur

---

**Conclusion:** Les modifications apportées à `generator.py` sont ✅ complètes et fonctionnelles. Les recommandations pour `convert.js` sont optionnelles mais très utiles pour diagnostiquer le manque de pièces jointes (de 70 attendues à 2 actuelles).
