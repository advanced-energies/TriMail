// taskpane.js
const SUBJECT_REGEX = /[PQME]\.\d{7}\.\d\.\d{2}/g;
const REST_URL = "https://outlook.office.com/api/v2.0"; // On n'utilise PAS graph.microsoft.com ici

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sortSelectedBtn").onclick = trierEmailSelectionne;
    document.getElementById("sortSentBtn").onclick = trierEmailsEnvoyes;
  }
});

// 1. Récupération du jeton interne (Zéro config Azure nécessaire)
async function getOutlookToken() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject("Erreur de jeton: " + JSON.stringify(result.error));
      }
    });
  });
}

// 2. Utilitaire d'appel API
async function callOutlookAPI(endpoint, method = "GET", body = null) {
  const token = await getOutlookToken();
  const url = endpoint.startsWith("http") ? endpoint : `${REST_URL}${endpoint}`;
  
  const response = await fetch(url, {
    method: method,
    headers: {
      "Authorization": `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: body ? JSON.stringify(body) : null
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`API Error ${response.status}: ${errorText}`);
  }
  return response.status === 204 ? {} : response.json();
}

// 3. Recherche récursive de dossier (Adaptée REST v2.0)
async function findFolder(startFolderId, fragment) {
  // On récupère les sous-dossiers
  const data = await callOutlookAPI(`/me/MailFolders/${startFolderId}/childfolders?$top=100`);
  
  for (const folder of data.value) {
    if (folder.DisplayName.toLowerCase().includes(fragment.toLowerCase())) {
      return folder.Id;
    }
    // Si le dossier a des enfants, on cherche dedans
    if (folder.ChildFolderCount > 0) {
      const subFolderId = await findFolder(folder.Id, fragment);
      if (subFolderId) return subFolderId;
    }
  }
  return null;
}

// 4. Action : Trier l'email sélectionné
async function trierEmailSelectionne() {
  try {
    const item = Office.context.mailbox.item;
    const matches = item.subject.match(SUBJECT_REGEX);
    
    if (!matches) {
      alert("Aucun motif [PQME].1234567.1.12 trouvé dans l'objet.");
      return;
    }

    const itemId = Office.context.mailbox.convertToRestId(
      item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );

    for (const code of matches) {
      console.log("Recherche du dossier pour :", code);
      const destId = await findFolder("inbox", code); // On commence la recherche dans la boîte de réception

      if (destId) {
        await callOutlookAPI(`/me/messages/${itemId}/move`, "POST", {
          "DestinationId": destId
        });
        alert(`Email déplacé avec succès vers : ${code}`);
      } else {
        alert(`Dossier introuvable pour le code : ${code}`);
      }
    }
  } catch (error) {
    console.error(error);
    alert("Erreur lors du tri : " + error.message);
  }
}

// 5. Action : Trier les éléments envoyés
async function trierEmailsEnvoyes() {
  try {
    alert("Lancement du tri des messages envoyés (10 derniers)...");
    const sentItems = await callOutlookAPI("/me/MailFolders/sentitems/messages?$top=10&$select=Subject,Id");

    for (const mail of sentItems.value) {
      const match = mail.Subject.match(SUBJECT_REGEX);
      if (match) {
        const destId = await findFolder("inbox", match[0]);
        if (destId) {
          await callOutlookAPI(`/me/messages/${mail.Id}/move`, "POST", {
            "DestinationId": destId
          });
          console.log(`Envoyé déplacé : ${mail.Subject}`);
        }
      }
    }
    alert("Tri des éléments envoyés terminé.");
  } catch (error) {
    console.error(error);
    alert("Erreur éléments envoyés : " + error.message);
  }
}