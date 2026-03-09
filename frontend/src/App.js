import "@/App.css";
import { Download, FileSpreadsheet, CheckSquare, Users, LogOut } from "lucide-react";

const Home = () => {
  const handleDownload = () => {
    const link = document.createElement('a');
    link.href = '/formulaire_rh_somnum.xlsx';
    link.download = 'formulaire_rh_somnum.xlsx';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-[#1B3A5F] to-[#0d1f33] flex items-center justify-center p-6">
      <div className="max-w-2xl w-full">
        {/* Logo */}
        <div className="text-center mb-8">
          <h1 className="text-4xl font-bold text-white mb-2">SOMNUM</h1>
          <p className="text-[#00A0B0] text-lg">Formulaire RH Informatique</p>
        </div>

        {/* Card principale */}
        <div className="bg-white/10 backdrop-blur-lg rounded-2xl p-8 border border-white/20">
          <div className="flex items-center gap-4 mb-6">
            <div className="bg-[#00A0B0] p-4 rounded-xl">
              <FileSpreadsheet className="w-10 h-10 text-white" />
            </div>
            <div>
              <h2 className="text-2xl font-semibold text-white">Formulaire Excel Amélioré</h2>
              <p className="text-gray-300">Version simplifiée avec menus déroulants</p>
            </div>
          </div>

          {/* Fonctionnalités */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-8">
            <div className="flex items-center gap-3 bg-white/5 rounded-lg p-4">
              <CheckSquare className="w-6 h-6 text-[#00A0B0]" />
              <span className="text-white">Menus déroulants</span>
            </div>
            <div className="flex items-center gap-3 bg-white/5 rounded-lg p-4">
              <Users className="w-6 h-6 text-[#00A0B0]" />
              <span className="text-white">Onglet Arrivée</span>
            </div>
            <div className="flex items-center gap-3 bg-white/5 rounded-lg p-4">
              <LogOut className="w-6 h-6 text-[#00A0B0]" />
              <span className="text-white">Onglet Départ</span>
            </div>
            <div className="flex items-center gap-3 bg-white/5 rounded-lg p-4">
              <FileSpreadsheet className="w-6 h-6 text-[#00A0B0]" />
              <span className="text-white">Guide d'utilisation</span>
            </div>
          </div>

          {/* Bouton téléchargement */}
          <button
            onClick={handleDownload}
            data-testid="download-btn"
            className="w-full bg-[#00A0B0] hover:bg-[#008a99] text-white font-semibold py-4 px-6 rounded-xl flex items-center justify-center gap-3 transition-all duration-300 transform hover:scale-[1.02] shadow-lg hover:shadow-[#00A0B0]/30"
          >
            <Download className="w-6 h-6" />
            Télécharger le formulaire Excel
          </button>

          <p className="text-center text-gray-400 text-sm mt-4">
            formulaire_rh_somnum.xlsx • Couleurs Somnum
          </p>
        </div>

        {/* Footer */}
        <p className="text-center text-gray-500 text-sm mt-6">
          © 2026 Somnum - Centres du sommeil
        </p>
      </div>
    </div>
  );
};

function App() {
  return <Home />;
}

export default App;
