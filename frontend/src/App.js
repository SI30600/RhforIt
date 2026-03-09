import { useState } from "react";
import "@/App.css";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import { 
  Download, FileSpreadsheet, User, Building2, Monitor, 
  Smartphone, Settings, LogOut, ChevronRight, Check,
  Calendar, Phone, Laptop, Camera, Printer, FolderOpen
} from "lucide-react";

// Couleurs Somnum
const COLORS = {
  bleuFonce: "#1B3A5F",
  turquoise: "#00A0B0",
  blanc: "#FFFFFF",
  grisClair: "#F5F7FA",
  grisMoyen: "#E8ECF0"
};

const CENTRES = [
  "Arles", "Nîmes", "Alès", "Montpellier", "Aubenas", 
  "Avignon", "Le Mans", "Lyon", "Firminy", "Rodez", "Aurillac", "Siège", "Autre"
];

const FONCTIONS = [
  "Médecin", "IDE", "Secrétaire", "Technicien(ne) du sommeil", 
  "Assistant(e) médical(e)", "Sage-femme", "Personnel Siège", "Autre"
];

const LOGICIELS = ["Nox", "Lowenstein", "Les deux", "Aucun"];
const TAILLES_PC = ['Petit (13-14")', 'Normal (15.6")', 'Grand (17")'];

function App() {
  const [activeTab, setActiveTab] = useState("arrivee");
  
  // État formulaire Arrivée
  const [arrivee, setArrivee] = useState({
    dateDemande: "",
    dateArrivee: "",
    nom: "",
    prenom: "",
    telephone: "",
    centre: "",
    centreAutre: "",
    statut: "",
    fonction: "",
    besoinOrdinateur: "",
    typeOrdinateur: "",
    taillePC: "",
    camera: "",
    besoinTelephone: "",
    logiciel: "",
    somnobook: "",
    imprimante: "",
    accesDrive: "",
    // Checklist IT
    balMicrosoft: false,
    intune: false,
    antivirus: false,
    drivePartage: false,
    signatureMail: false,
    entraId: false,
    o365Chrome: false
  });

  // État formulaire Départ
  const [depart, setDepart] = useState({
    dateDemande: "",
    dateDepart: "",
    nom: "",
    prenom: "",
    centre: "",
    centreAutre: "",
    sauvegardeMail: false,
    suppressionCompte: false,
    liberationLicence: false,
    recuperationMateriel: false,
    desactivationDrive: false,
    suppressionIntune: false,
    commentaires: ""
  });

  const handleArriveeChange = (field, value) => {
    setArrivee(prev => ({ ...prev, [field]: value }));
  };

  const handleDepartChange = (field, value) => {
    setDepart(prev => ({ ...prev, [field]: value }));
  };

  const exportToExcel = () => {
    const wb = XLSX.utils.book_new();

    // Feuille Arrivée
    const arriveeData = [
      ["FORMULAIRE D'ARRIVÉE - SOMNUM"],
      [],
      ["INFORMATIONS DU COLLABORATEUR"],
      ["Date de la demande", arrivee.dateDemande],
      ["Date d'arrivée prévue", arrivee.dateArrivee],
      ["Nom", arrivee.nom],
      ["Prénom", arrivee.prenom],
      ["Téléphone personnel", arrivee.telephone],
      ["Centre de rattachement", arrivee.centre === "Autre" ? arrivee.centreAutre : arrivee.centre],
      [],
      ["TYPE DE PERSONNEL"],
      ["Statut", arrivee.statut],
      ["Fonction", arrivee.fonction],
      [],
      ["BESOINS INFORMATIQUES"],
      ["Besoin ordinateur", arrivee.besoinOrdinateur],
      ["Type d'ordinateur", arrivee.typeOrdinateur],
      ["Taille PC portable", arrivee.taillePC],
      ["Caméra intégrée", arrivee.camera],
      ["Besoin téléphone pro", arrivee.besoinTelephone],
      ["Logiciel métier", arrivee.logiciel],
      ["Raccourci SomnoBook", arrivee.somnobook],
      ["Ajout imprimante", arrivee.imprimante],
      ["Accès sites Drive", arrivee.accesDrive],
      [],
      ["CONFIGURATION IT"],
      ["Création BAL Microsoft", arrivee.balMicrosoft ? "Fait" : "À faire"],
      ["Intégration Intune", arrivee.intune ? "Fait" : "À faire"],
      ["Installation Antivirus", arrivee.antivirus ? "Fait" : "À faire"],
      ["Configuration Drive partagé", arrivee.drivePartage ? "Fait" : "À faire"],
      ["Création signature mail", arrivee.signatureMail ? "Fait" : "À faire"],
      ["Microsoft Entra ID", arrivee.entraId ? "Fait" : "À faire"],
      ["Installation O365 + Chrome", arrivee.o365Chrome ? "Fait" : "À faire"]
    ];
    const wsArrivee = XLSX.utils.aoa_to_sheet(arriveeData);
    wsArrivee["!cols"] = [{ wch: 30 }, { wch: 40 }];
    XLSX.utils.book_append_sheet(wb, wsArrivee, "ARRIVÉE");

    // Feuille Départ
    const departData = [
      ["FORMULAIRE DE DÉPART - SOMNUM"],
      [],
      ["INFORMATIONS DU COLLABORATEUR"],
      ["Date de la demande", depart.dateDemande],
      ["Date de départ", depart.dateDepart],
      ["Nom", depart.nom],
      ["Prénom", depart.prenom],
      ["Centre de rattachement", depart.centre === "Autre" ? depart.centreAutre : depart.centre],
      [],
      ["ACTIONS À RÉALISER"],
      ["Sauvegarde des mails", depart.sauvegardeMail ? "Fait" : "À faire"],
      ["Suppression compte Microsoft", depart.suppressionCompte ? "Fait" : "À faire"],
      ["Libération licence Microsoft", depart.liberationLicence ? "Fait" : "À faire"],
      ["Récupération matériel", depart.recuperationMateriel ? "Fait" : "À faire"],
      ["Désactivation accès Drive", depart.desactivationDrive ? "Fait" : "À faire"],
      ["Suppression Intune", depart.suppressionIntune ? "Fait" : "À faire"],
      [],
      ["Commentaires", depart.commentaires]
    ];
    const wsDepart = XLSX.utils.aoa_to_sheet(departData);
    wsDepart["!cols"] = [{ wch: 30 }, { wch: 40 }];
    XLSX.utils.book_append_sheet(wb, wsDepart, "DÉPART");

    // Export
    const excelBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const data = new Blob([excelBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(data, "formulaire_rh_somnum.xlsx");
  };

  // Composants réutilisables
  const SelectField = ({ label, value, onChange, options, icon: Icon, placeholder = "Sélectionner..." }) => (
    <div className="space-y-2">
      <label className="flex items-center gap-2 text-sm font-medium text-gray-300">
        {Icon && <Icon className="w-4 h-4 text-[#00A0B0]" />}
        {label}
      </label>
      <select
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className="w-full bg-white/10 border border-white/20 rounded-lg px-4 py-3 text-white focus:outline-none focus:ring-2 focus:ring-[#00A0B0] focus:border-transparent transition-all cursor-pointer"
        data-testid={`select-${label.toLowerCase().replace(/\s/g, '-')}`}
      >
        <option value="" className="bg-[#1B3A5F]">{placeholder}</option>
        {options.map(opt => (
          <option key={opt} value={opt} className="bg-[#1B3A5F]">{opt}</option>
        ))}
      </select>
    </div>
  );

  const TextField = ({ label, value, onChange, icon: Icon, placeholder = "", type = "text" }) => (
    <div className="space-y-2">
      <label className="flex items-center gap-2 text-sm font-medium text-gray-300">
        {Icon && <Icon className="w-4 h-4 text-[#00A0B0]" />}
        {label}
      </label>
      <input
        type={type}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        className="w-full bg-white/10 border border-white/20 rounded-lg px-4 py-3 text-white placeholder-gray-500 focus:outline-none focus:ring-2 focus:ring-[#00A0B0] focus:border-transparent transition-all"
        data-testid={`input-${label.toLowerCase().replace(/\s/g, '-')}`}
      />
    </div>
  );

  const Checkbox = ({ label, checked, onChange }) => (
    <label className="flex items-center gap-3 cursor-pointer group" data-testid={`checkbox-${label.toLowerCase().replace(/\s/g, '-')}`}>
      <div 
        className={`w-6 h-6 rounded-md border-2 flex items-center justify-center transition-all ${
          checked 
            ? 'bg-[#00A0B0] border-[#00A0B0]' 
            : 'border-white/30 group-hover:border-[#00A0B0]'
        }`}
        onClick={() => onChange(!checked)}
      >
        {checked && <Check className="w-4 h-4 text-white" />}
      </div>
      <span className="text-white text-sm">{label}</span>
    </label>
  );

  const SectionTitle = ({ children, icon: Icon }) => (
    <div className="flex items-center gap-3 mb-4">
      <div className="bg-[#00A0B0] p-2 rounded-lg">
        <Icon className="w-5 h-5 text-white" />
      </div>
      <h3 className="text-lg font-semibold text-white">{children}</h3>
    </div>
  );

  return (
    <div className="min-h-screen bg-gradient-to-br from-[#1B3A5F] to-[#0d1f33]">
      {/* Header */}
      <header className="bg-[#1B3A5F]/80 backdrop-blur-lg border-b border-white/10 sticky top-0 z-50">
        <div className="max-w-6xl mx-auto px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <FileSpreadsheet className="w-8 h-8 text-[#00A0B0]" />
            <div>
              <h1 className="text-xl font-bold text-white">SOMNUM</h1>
              <p className="text-xs text-[#00A0B0]">Formulaire RH Informatique</p>
            </div>
          </div>
          <button
            onClick={exportToExcel}
            data-testid="export-btn"
            className="bg-[#00A0B0] hover:bg-[#008a99] text-white font-medium py-2.5 px-5 rounded-lg flex items-center gap-2 transition-all transform hover:scale-105 shadow-lg"
          >
            <Download className="w-5 h-5" />
            Exporter Excel
          </button>
        </div>
      </header>

      {/* Tabs */}
      <div className="max-w-6xl mx-auto px-6 pt-6">
        <div className="flex gap-2">
          <button
            onClick={() => setActiveTab("arrivee")}
            data-testid="tab-arrivee"
            className={`flex items-center gap-2 px-6 py-3 rounded-t-xl font-medium transition-all ${
              activeTab === "arrivee"
                ? "bg-white/10 text-white border-t-2 border-[#00A0B0]"
                : "text-gray-400 hover:text-white hover:bg-white/5"
            }`}
          >
            <User className="w-5 h-5" />
            Arrivée
          </button>
          <button
            onClick={() => setActiveTab("depart")}
            data-testid="tab-depart"
            className={`flex items-center gap-2 px-6 py-3 rounded-t-xl font-medium transition-all ${
              activeTab === "depart"
                ? "bg-white/10 text-white border-t-2 border-[#00A0B0]"
                : "text-gray-400 hover:text-white hover:bg-white/5"
            }`}
          >
            <LogOut className="w-5 h-5" />
            Départ
          </button>
        </div>
      </div>

      {/* Content */}
      <div className="max-w-6xl mx-auto px-6 pb-10">
        <div className="bg-white/5 backdrop-blur-lg rounded-b-2xl rounded-tr-2xl border border-white/10 p-8">
          
          {/* FORMULAIRE ARRIVÉE */}
          {activeTab === "arrivee" && (
            <div className="space-y-8">
              {/* Section Informations */}
              <section>
                <SectionTitle icon={User}>Informations du collaborateur</SectionTitle>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                  <TextField 
                    label="Date de la demande" 
                    value={arrivee.dateDemande} 
                    onChange={(v) => handleArriveeChange("dateDemande", v)}
                    icon={Calendar}
                    type="date"
                  />
                  <TextField 
                    label="Date d'arrivée prévue" 
                    value={arrivee.dateArrivee} 
                    onChange={(v) => handleArriveeChange("dateArrivee", v)}
                    icon={Calendar}
                    type="date"
                  />
                  <SelectField 
                    label="Centre de rattachement" 
                    value={arrivee.centre} 
                    onChange={(v) => handleArriveeChange("centre", v)}
                    options={CENTRES}
                    icon={Building2}
                  />
                  {arrivee.centre === "Autre" && (
                    <TextField 
                      label="Préciser le centre" 
                      value={arrivee.centreAutre} 
                      onChange={(v) => handleArriveeChange("centreAutre", v)}
                      icon={Building2}
                      placeholder="Nom du centre..."
                    />
                  )}
                  <TextField 
                    label="Nom" 
                    value={arrivee.nom} 
                    onChange={(v) => handleArriveeChange("nom", v)}
                    icon={User}
                  />
                  <TextField 
                    label="Prénom" 
                    value={arrivee.prenom} 
                    onChange={(v) => handleArriveeChange("prenom", v)}
                    icon={User}
                  />
                  <TextField 
                    label="Téléphone personnel" 
                    value={arrivee.telephone} 
                    onChange={(v) => handleArriveeChange("telephone", v)}
                    icon={Phone}
                    placeholder="Pour création compte Google"
                  />
                </div>
              </section>

              {/* Section Type Personnel */}
              <section>
                <SectionTitle icon={Building2}>Type de personnel</SectionTitle>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <SelectField 
                    label="Statut" 
                    value={arrivee.statut} 
                    onChange={(v) => handleArriveeChange("statut", v)}
                    options={["Salarié", "Collaborateur"]}
                  />
                  <SelectField 
                    label="Fonction" 
                    value={arrivee.fonction} 
                    onChange={(v) => handleArriveeChange("fonction", v)}
                    options={FONCTIONS}
                  />
                </div>
              </section>

              {/* Section Besoins Informatiques */}
              <section>
                <SectionTitle icon={Monitor}>Besoins informatiques</SectionTitle>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                  <SelectField 
                    label="Besoin d'un ordinateur" 
                    value={arrivee.besoinOrdinateur} 
                    onChange={(v) => handleArriveeChange("besoinOrdinateur", v)}
                    options={["Oui", "Non"]}
                    icon={Monitor}
                  />
                  <SelectField 
                    label="Type d'ordinateur" 
                    value={arrivee.typeOrdinateur} 
                    onChange={(v) => handleArriveeChange("typeOrdinateur", v)}
                    options={["PC Fixe", "PC Portable"]}
                    icon={Laptop}
                  />
                  <SelectField 
                    label="Taille PC portable" 
                    value={arrivee.taillePC} 
                    onChange={(v) => handleArriveeChange("taillePC", v)}
                    options={TAILLES_PC}
                  />
                  <SelectField 
                    label="Caméra intégrée" 
                    value={arrivee.camera} 
                    onChange={(v) => handleArriveeChange("camera", v)}
                    options={["Oui", "Non"]}
                    icon={Camera}
                  />
                  <SelectField 
                    label="Besoin téléphone pro" 
                    value={arrivee.besoinTelephone} 
                    onChange={(v) => handleArriveeChange("besoinTelephone", v)}
                    options={["Oui", "Non"]}
                    icon={Smartphone}
                  />
                  <SelectField 
                    label="Logiciel métier" 
                    value={arrivee.logiciel} 
                    onChange={(v) => handleArriveeChange("logiciel", v)}
                    options={LOGICIELS}
                    icon={Settings}
                  />
                  <SelectField 
                    label="Raccourci SomnoBook" 
                    value={arrivee.somnobook} 
                    onChange={(v) => handleArriveeChange("somnobook", v)}
                    options={["Oui", "Non"]}
                  />
                  <SelectField 
                    label="Ajout imprimante" 
                    value={arrivee.imprimante} 
                    onChange={(v) => handleArriveeChange("imprimante", v)}
                    options={["Oui", "Non"]}
                    icon={Printer}
                  />
                  <TextField 
                    label="Accès sites Drive" 
                    value={arrivee.accesDrive} 
                    onChange={(v) => handleArriveeChange("accesDrive", v)}
                    icon={FolderOpen}
                    placeholder="Préciser les sites..."
                  />
                </div>
              </section>

              {/* Section Configuration IT */}
              <section>
                <SectionTitle icon={Settings}>Configuration IT (cocher quand fait)</SectionTitle>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 bg-white/5 rounded-xl p-6">
                  <Checkbox 
                    label="Création BAL Microsoft" 
                    checked={arrivee.balMicrosoft}
                    onChange={(v) => handleArriveeChange("balMicrosoft", v)}
                  />
                  <Checkbox 
                    label="Intégration sur Intune" 
                    checked={arrivee.intune}
                    onChange={(v) => handleArriveeChange("intune", v)}
                  />
                  <Checkbox 
                    label="Installation Antivirus" 
                    checked={arrivee.antivirus}
                    onChange={(v) => handleArriveeChange("antivirus", v)}
                  />
                  <Checkbox 
                    label="Configuration Drive partagé" 
                    checked={arrivee.drivePartage}
                    onChange={(v) => handleArriveeChange("drivePartage", v)}
                  />
                  <Checkbox 
                    label="Création signature mail" 
                    checked={arrivee.signatureMail}
                    onChange={(v) => handleArriveeChange("signatureMail", v)}
                  />
                  <Checkbox 
                    label="Microsoft Entra ID" 
                    checked={arrivee.entraId}
                    onChange={(v) => handleArriveeChange("entraId", v)}
                  />
                  <Checkbox 
                    label="Installation O365 + Chrome" 
                    checked={arrivee.o365Chrome}
                    onChange={(v) => handleArriveeChange("o365Chrome", v)}
                  />
                </div>
              </section>
            </div>
          )}

          {/* FORMULAIRE DÉPART */}
          {activeTab === "depart" && (
            <div className="space-y-8">
              {/* Section Informations */}
              <section>
                <SectionTitle icon={User}>Informations du collaborateur</SectionTitle>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                  <TextField 
                    label="Date de la demande" 
                    value={depart.dateDemande} 
                    onChange={(v) => handleDepartChange("dateDemande", v)}
                    icon={Calendar}
                    type="date"
                  />
                  <TextField 
                    label="Date de départ" 
                    value={depart.dateDepart} 
                    onChange={(v) => handleDepartChange("dateDepart", v)}
                    icon={Calendar}
                    type="date"
                  />
                  <SelectField 
                    label="Centre de rattachement" 
                    value={depart.centre} 
                    onChange={(v) => handleDepartChange("centre", v)}
                    options={CENTRES}
                    icon={Building2}
                  />
                  {depart.centre === "Autre" && (
                    <TextField 
                      label="Préciser le centre" 
                      value={depart.centreAutre} 
                      onChange={(v) => handleDepartChange("centreAutre", v)}
                      icon={Building2}
                      placeholder="Nom du centre..."
                    />
                  )}
                  <TextField 
                    label="Nom" 
                    value={depart.nom} 
                    onChange={(v) => handleDepartChange("nom", v)}
                    icon={User}
                  />
                  <TextField 
                    label="Prénom" 
                    value={depart.prenom} 
                    onChange={(v) => handleDepartChange("prenom", v)}
                    icon={User}
                  />
                </div>
              </section>

              {/* Section Actions Départ */}
              <section>
                <SectionTitle icon={Settings}>Actions à réaliser (cocher quand fait)</SectionTitle>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 bg-white/5 rounded-xl p-6">
                  <Checkbox 
                    label="Sauvegarde des mails (OneDrive KENNEDY\SIEGE\IT)" 
                    checked={depart.sauvegardeMail}
                    onChange={(v) => handleDepartChange("sauvegardeMail", v)}
                  />
                  <Checkbox 
                    label="Suppression compte Microsoft" 
                    checked={depart.suppressionCompte}
                    onChange={(v) => handleDepartChange("suppressionCompte", v)}
                  />
                  <Checkbox 
                    label="Libération licence Microsoft" 
                    checked={depart.liberationLicence}
                    onChange={(v) => handleDepartChange("liberationLicence", v)}
                  />
                  <Checkbox 
                    label="Récupération matériel (PC / Téléphone)" 
                    checked={depart.recuperationMateriel}
                    onChange={(v) => handleDepartChange("recuperationMateriel", v)}
                  />
                  <Checkbox 
                    label="Désactivation accès Drive" 
                    checked={depart.desactivationDrive}
                    onChange={(v) => handleDepartChange("desactivationDrive", v)}
                  />
                  <Checkbox 
                    label="Suppression Intune" 
                    checked={depart.suppressionIntune}
                    onChange={(v) => handleDepartChange("suppressionIntune", v)}
                  />
                </div>
              </section>

              {/* Commentaires */}
              <section>
                <SectionTitle icon={FileSpreadsheet}>Commentaires</SectionTitle>
                <textarea
                  value={depart.commentaires}
                  onChange={(e) => handleDepartChange("commentaires", e.target.value)}
                  placeholder="Ajouter des commentaires..."
                  className="w-full bg-white/10 border border-white/20 rounded-xl px-4 py-3 text-white placeholder-gray-500 focus:outline-none focus:ring-2 focus:ring-[#00A0B0] focus:border-transparent transition-all min-h-[120px] resize-none"
                  data-testid="textarea-commentaires"
                />
              </section>
            </div>
          )}
        </div>
      </div>

      {/* Footer */}
      <footer className="text-center py-6 text-gray-500 text-sm">
        © 2026 Somnum - Centres du sommeil
      </footer>
    </div>
  );
}

export default App;
