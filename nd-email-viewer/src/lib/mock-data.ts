import { NDCabinet, NDDocument, NDFolder, NDWorkspace } from "./types";

// Realistic law firm email and document data
const people = [
  "Andrew Meston",
  "Sarah Mitchell",
  "James Robertson",
  "Claire Thompson",
  "David Wallace",
  "Emma Chambers",
  "Michael O'Brien",
  "Rachel Sinclair",
  "Thomas Whitfield",
  "Laura Campbell",
];

const emailAddresses: Record<string, string> = {
  "Andrew Meston": "andrew.meston@wallacefirm.com",
  "Sarah Mitchell": "sarah.mitchell@wallacefirm.com",
  "James Robertson": "james.robertson@oppco.com",
  "Claire Thompson": "claire.thompson@wallacefirm.com",
  "David Wallace": "david.wallace@wallacefirm.com",
  "Emma Chambers": "emma.chambers@clientcorp.com",
  "Michael O'Brien": "michael.obrien@wallacefirm.com",
  "Rachel Sinclair": "rachel.sinclair@chambers.co.uk",
  "Thomas Whitfield": "thomas.whitfield@oppco.com",
  "Laura Campbell": "laura.campbell@wallacefirm.com",
};

const emailSubjects = [
  "RE: Settlement offer - revised terms",
  "FW: Disclosure bundle - missing documents",
  "RE: Court hearing - 14 March",
  "Expert report - Dr Henderson",
  "RE: Costs schedule draft",
  "FW: Witness statement - final version",
  "RE: Mediation date confirmation",
  "Without prejudice - settlement discussion",
  "RE: Bundle index - agreed version",
  "FW: Counsel's advice on quantum",
  "RE: Pre-trial checklist",
  "Joint statement of experts",
  "RE: Application for extension of time",
  "FW: Updated schedule of loss",
  "RE: Part 36 offer - response deadline",
  "Skeleton argument - first draft",
  "RE: Case management conference prep",
  "FW: Insurance policy documents",
  "RE: Chronology - amendments needed",
  "Directions order - sealed copy",
  "RE: Funding agreement review",
  "FW: Medical records request",
  "RE: Trial window - availability",
  "Agreed bundle of authorities",
  "RE: Disclosure pilot - document review",
];

const docNames = [
  "Particulars of Claim - FINAL.docx",
  "Defence and Counterclaim v3.docx",
  "Witness Statement - David Wallace.docx",
  "Expert Report - Quantum - Dr Henderson.pdf",
  "Schedule of Loss - Updated Feb 2026.xlsx",
  "Costs Budget - Approved.xlsx",
  "Counsel Fee Note - March 2026.pdf",
  "Chronology - Master.docx",
  "Bundle Index - Agreed.docx",
  "Attendance Note - CMC 12 Jan 2026.docx",
  "Disclosure List - Claimant.xlsx",
  "Trial Bundle - Section A.pdf",
  "Settlement Agreement - Draft.docx",
  "Part 36 Offer Letter.pdf",
  "Client Care Letter - Signed.pdf",
  "CFA Agreement - Executed.pdf",
  "Court Order - Directions 15 Dec 2025.pdf",
  "Skeleton Argument - Trial.docx",
  "List of Issues - Agreed.docx",
  "Pre-Trial Checklist.pdf",
];

function randomDate(start: Date, end: Date): string {
  const d = new Date(
    start.getTime() + Math.random() * (end.getTime() - start.getTime())
  );
  return d.toISOString();
}

function randomItem<T>(arr: T[]): T {
  return arr[Math.floor(Math.random() * arr.length)];
}

function randomSize(): number {
  return Math.floor(Math.random() * 5000000) + 10000;
}

// Generate emails with deliberately jumbled dates per sender (to demonstrate the problem)
function generateEmails(count: number, folderId: string): NDDocument[] {
  const emails: NDDocument[] = [];
  // Create clusters of emails from same senders with non-sequential dates
  const senderGroups = [
    { sender: "Andrew Meston", count: 10 },
    { sender: "Sarah Mitchell", count: 8 },
    { sender: "James Robertson", count: 6 },
    { sender: "Claire Thompson", count: 5 },
    { sender: "Emma Chambers", count: 4 },
    { sender: "Rachel Sinclair", count: 3 },
    { sender: "Thomas Whitfield", count: 3 },
    { sender: "David Wallace", count: 2 },
  ];

  let id = 1;
  for (const group of senderGroups) {
    for (let i = 0; i < group.count; i++) {
      const sentDate = randomDate(
        new Date("2025-06-01"),
        new Date("2026-03-25")
      );
      const recipients = [randomItem(people.filter((p) => p !== group.sender))];
      if (Math.random() > 0.6)
        recipients.push(
          randomItem(people.filter((p) => !recipients.includes(p) && p !== group.sender))
        );

      emails.push({
        id: `${folderId}-email-${id++}`,
        name: randomItem(emailSubjects),
        type: "email",
        extension: "msg",
        size: randomSize(),
        createdDate: sentDate,
        modifiedDate: sentDate,
        createdBy: group.sender,
        modifiedBy: group.sender,
        version: 1,
        emailAttributes: {
          from: `${group.sender} <${emailAddresses[group.sender]}>`,
          to: recipients.map((r) => `${r} <${emailAddresses[r]}>`),
          cc: Math.random() > 0.7
            ? [randomItem(people.filter((p) => !recipients.includes(p) && p !== group.sender))].map(
                (r) => `${r} <${emailAddresses[r]}>`
              )
            : undefined,
          sentDate,
          subject: randomItem(emailSubjects),
        },
        client: "ClientCorp Ltd",
        matter: "Robertson v ClientCorp",
      });
    }
  }

  // Shuffle to simulate ND's random ordering
  return emails.sort(() => Math.random() - 0.5);
}

function generateDocuments(count: number, folderId: string): NDDocument[] {
  const docs: NDDocument[] = [];
  for (let i = 0; i < count; i++) {
    const name = i < docNames.length ? docNames[i] : `Document ${i + 1}.docx`;
    const ext = name.split(".").pop() || "docx";
    const created = randomDate(new Date("2025-01-01"), new Date("2026-03-25"));
    const modified = randomDate(new Date(created), new Date("2026-03-25"));

    docs.push({
      id: `${folderId}-doc-${i + 1}`,
      name,
      type: "document",
      extension: ext,
      size: randomSize(),
      createdDate: created,
      modifiedDate: modified,
      createdBy: randomItem(people),
      modifiedBy: randomItem(people),
      version: Math.floor(Math.random() * 5) + 1,
      client: "ClientCorp Ltd",
      matter: "Robertson v ClientCorp",
    });
  }
  return docs;
}

// Build folder tree
const emailFolders: NDFolder[] = [
  {
    id: "folder-emails-main",
    name: "Correspondence",
    parentId: "ws-1",
    folderType: "emails",
    documentCount: 41,
  },
  {
    id: "folder-emails-counsel",
    name: "Counsel Correspondence",
    parentId: "ws-1",
    folderType: "emails",
    documentCount: 15,
  },
  {
    id: "folder-emails-expert",
    name: "Expert Correspondence",
    parentId: "ws-1",
    folderType: "emails",
    documentCount: 8,
  },
  {
    id: "folder-emails-client",
    name: "Client Correspondence",
    parentId: "ws-1",
    folderType: "emails",
    documentCount: 12,
  },
];

const docFolders: NDFolder[] = [
  {
    id: "folder-docs-pleadings",
    name: "Pleadings",
    parentId: "ws-1",
    folderType: "documents",
    documentCount: 8,
  },
  {
    id: "folder-docs-evidence",
    name: "Evidence",
    parentId: "ws-1",
    folderType: "documents",
    documentCount: 12,
  },
  {
    id: "folder-docs-costs",
    name: "Costs",
    parentId: "ws-1",
    folderType: "documents",
    documentCount: 6,
  },
  {
    id: "folder-docs-admin",
    name: "Administration",
    parentId: "ws-1",
    folderType: "documents",
    documentCount: 5,
  },
];

const workspace: NDWorkspace = {
  id: "ws-1",
  name: "Robertson v ClientCorp Ltd",
  client: "ClientCorp Ltd",
  matter: "MAT-2025-0142",
  folders: [...emailFolders, ...docFolders],
};

const workspace2: NDWorkspace = {
  id: "ws-2",
  name: "Thompson Property Acquisition",
  client: "Thompson Holdings",
  matter: "MAT-2025-0287",
  folders: [
    {
      id: "folder-w2-emails",
      name: "Correspondence",
      parentId: "ws-2",
      folderType: "emails",
      documentCount: 18,
    },
    {
      id: "folder-w2-docs",
      name: "Transaction Documents",
      parentId: "ws-2",
      folderType: "documents",
      documentCount: 14,
    },
    {
      id: "folder-w2-dd",
      name: "Due Diligence",
      parentId: "ws-2",
      folderType: "documents",
      documentCount: 22,
    },
  ],
};

export const mockCabinet: NDCabinet = {
  id: "cab-1",
  name: "Wallace & Partners",
  workspaces: [workspace, workspace2],
};

// Document store - keyed by folder ID
const documentStore: Record<string, NDDocument[]> = {};

export function getFolderDocuments(folderId: string): NDDocument[] {
  if (!documentStore[folderId]) {
    const allFolders = [...mockCabinet.workspaces.flatMap((ws) => ws.folders)];
    const folder = allFolders.find((f) => f.id === folderId);
    if (!folder) return [];

    if (folder.folderType === "emails") {
      documentStore[folderId] = generateEmails(folder.documentCount, folderId);
    } else {
      documentStore[folderId] = generateDocuments(
        folder.documentCount,
        folderId
      );
    }
  }
  return documentStore[folderId];
}

export function addDocumentToFolder(
  folderId: string,
  doc: NDDocument
): NDDocument[] {
  if (!documentStore[folderId]) {
    getFolderDocuments(folderId); // Initialize
  }
  documentStore[folderId] = [doc, ...documentStore[folderId]];
  return documentStore[folderId];
}

export function removeDocumentFromFolder(
  folderId: string,
  docId: string
): NDDocument[] {
  if (!documentStore[folderId]) return [];
  documentStore[folderId] = documentStore[folderId].filter(
    (d) => d.id !== docId
  );
  return documentStore[folderId];
}
