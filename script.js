// 1. DATA CONFIGURATION
const toolsData = [
    // ORGANIZE PDF
    { category: "Organize PDF", icon: "ph-files", color: "text-red-500", bg: "bg-red-50", title: "Merge PDF", desc: "Combine PDFs in the order you want with the easiest PDF merger available.", link: "MergePDF.html" },
    { category: "Organize PDF", icon: "ph-scissors", color: "text-red-500", bg: "bg-red-50", title: "Split PDF", desc: "Separate one page or a whole set for easy conversion into independent PDF files.", link: "SplitPDF.html" },
    { category: "Organize PDF", icon: "ph-sort-ascending", color: "text-red-500", bg: "bg-red-50", title: "Organize PDF", desc: "Sort pages of your PDF file however you like. Delete or add PDF pages.", link: "#" },
    { category: "Organize PDF", icon: "ph-scan", color: "text-red-500", bg: "bg-red-50", title: "Scan to PDF", desc: "Capture document scans from your mobile device and send them instantly to your browser.", link: "#" },

    // OPTIMIZE PDF
    { category: "Optimize PDF", icon: "ph-arrows-in-line-horizontal", color: "text-green-500", bg: "bg-green-50", title: "Compress PDF", desc: "Reduce file size while optimizing for maximal PDF quality.", link: "CompressPDF.html" },
    { category: "Optimize PDF", icon: "ph-wrench", color: "text-green-500", bg: "bg-green-50", title: "Repair PDF", desc: "Repair a damaged PDF and recover data from corrupt PDF. Fix PDF files.", link: "#" },
    { category: "Optimize PDF", icon: "ph-text-aa", color: "text-green-500", bg: "bg-green-50", title: "OCR PDF", desc: "Easily convert scanned PDF into searchable and selectable documents.", link: "#" },

    // CONVERT PDF
    { category: "Convert PDF", icon: "ph-file-doc", color: "text-blue-500", bg: "bg-blue-50", title: "PDF to Word", desc: "Easily convert your PDF files into easy to edit DOC and DOCX documents.", link: "PDFToWord.html" },
    { category: "Convert PDF", icon: "ph-presentation", color: "text-red-500", bg: "bg-red-50", title: "PDF to PowerPoint", desc: "Turn your PDF files into easy to edit PPT and PPTX slideshows.", link: "PDFtoPPT.html" },
    { category: "Convert PDF", icon: "ph-file-xls", color: "text-green-500", bg: "bg-green-50", title: "PDF to Excel", desc: "Pull data straight from PDFs into Excel spreadsheets in a few short seconds.", link: "PDFtoExcel.html" },
    { category: "Convert PDF", icon: "ph-microsoft-word-logo", color: "text-blue-500", bg: "bg-blue-50", title: "Word to PDF", desc: "Make DOC and DOCX files easy to read by converting them to PDF.", link: "#" },
    { category: "Convert PDF", icon: "ph-microsoft-powerpoint-logo", color: "text-orange-500", bg: "bg-orange-50", title: "PowerPoint to PDF", desc: "Make PPT and PPTX slideshows easy to view by converting them to PDF.", link: "#" },
    { category: "Convert PDF", icon: "ph-microsoft-excel-logo", color: "text-green-500", bg: "bg-green-50", title: "Excel to PDF", desc: "Make EXCEL spreadsheets easy to read by converting them to PDF.", link: "#" },
    { category: "Convert PDF", icon: "ph-image", color: "text-yellow-500", bg: "bg-yellow-50", title: "PDF to JPG", desc: "Convert each PDF page into a JPG or extract all images contained in a PDF.", link: "#" },
    { category: "Convert PDF", icon: "ph-image-square", color: "text-yellow-500", bg: "bg-yellow-50", title: "JPG to PDF", desc: "Convert JPG images to PDF in seconds. Easily adjust orientation and margins.", link: "#" },
    { category: "Convert PDF", icon: "ph-code", color: "text-gray-500", bg: "bg-gray-50", title: "HTML to PDF", desc: "Convert webpages in HTML to PDF.", link: "#" },
    { category: "Convert PDF", icon: "ph-file-pdf", color: "text-blue-900", bg: "bg-blue-100", title: "PDF to PDF/A", desc: "Transform your PDF to PDF/A, the ISO-standardized version of PDF.", link: "#" },

    // EDIT PDF
    { category: "Edit PDF", icon: "ph-pen-nib", color: "text-purple-500", bg: "bg-purple-50", title: "Edit PDF", desc: "Add text, images, shapes or freehand annotations to a PDF document.", badge: "New!", link: "#" },
    { category: "Edit PDF", icon: "ph-stamp", color: "text-purple-500", bg: "bg-purple-50", title: "Watermark", desc: "Stamp an image or text over your PDF in seconds.", link: "#" },
    { category: "Edit PDF", icon: "ph-arrows-clockwise", color: "text-purple-500", bg: "bg-purple-50", title: "Rotate PDF", desc: "Rotate your PDFs the way you need them.", link: "#" },
    { category: "Edit PDF", icon: "ph-list-numbers", color: "text-purple-500", bg: "bg-purple-50", title: "Page numbers", desc: "Add page numbers into PDFs with ease.", link: "#" },
    { category: "Edit PDF", icon: "ph-crop", color: "text-purple-500", bg: "bg-purple-50", title: "Crop PDF", desc: "Crop margins of PDF documents or select specific areas.", badge: "New!", link: "#" },

    // PDF SECURITY
    { category: "PDF Security", icon: "ph-pen", color: "text-blue-500", bg: "bg-blue-50", title: "Sign PDF", desc: "Sign yourself or request electronic signatures from others.", link: "#" },
    { category: "PDF Security", icon: "ph-lock-key-open", color: "text-gray-500", bg: "bg-gray-50", title: "Unlock PDF", desc: "Remove PDF password security, giving you the freedom to use your PDFs as you want.", link: "#" },
    { category: "PDF Security", icon: "ph-shield-check", color: "text-blue-500", bg: "bg-blue-50", title: "Protect PDF", desc: "Protect PDF files with a password. Encrypt PDF documents.", link: "#" },
    { category: "PDF Security", icon: "ph-files", color: "text-blue-500", bg: "bg-blue-50", title: "Compare PDF", desc: "Show a side-by-side document comparison and easily spot changes.", badge: "New!", link: "#" },
    { category: "PDF Security", icon: "ph-eye-slash", color: "text-blue-500", bg: "bg-blue-50", title: "Redact PDF", desc: "Redact text and graphics to permanently remove sensitive information from a PDF.", badge: "New!", link: "#" }
];

const categoriesData = ["All", "Organize PDF", "Optimize PDF", "Convert PDF", "Edit PDF", "PDF Security"];

// Global state
let currentCategory = "All";

// 2. RENDER FUNCTIONS
function renderTools() {
    const grid = document.getElementById('tools-grid');
    let html = '';
    
    // Filter logic
    const filteredTools = currentCategory === "All" 
        ? toolsData 
        : toolsData.filter(tool => tool.category === currentCategory);

    filteredTools.forEach((tool, index) => {
        // Add staggered animation delay
        const delay = index * 50; 
        
        html += `
        <a href="${tool.link}" class="tool-card animate-fade-in bg-white dark:bg-slate-800 p-6 rounded-xl border border-slate-100 dark:border-slate-700 shadow-sm cursor-pointer group flex flex-col items-start text-left h-full relative overflow-hidden hover:no-underline" style="animation-delay: ${delay}ms">
            ${tool.badge ? `<span class="absolute top-4 right-4 text-red-500 bg-red-50 dark:bg-red-900/30 border border-red-100 dark:border-red-900/50 text-[10px] font-bold px-2 py-0.5 rounded uppercase tracking-wider">${tool.badge}</span>` : ''}
            
            <div class="mb-5 p-3 rounded-lg ${tool.bg} dark:bg-slate-700 ${tool.color} w-fit transform group-hover:scale-110 transition-transform duration-300">
                <i class="ph-fill ${tool.icon} text-3xl"></i>
            </div>
            <h3 class="text-lg font-bold text-slate-800 dark:text-slate-100 mb-2 group-hover:text-brand-600 dark:group-hover:text-brand-400 transition-colors">
                ${tool.title}
            </h3>
            <p class="text-slate-500 dark:text-slate-400 text-sm leading-relaxed">
                ${tool.desc}
            </p>
        </a>
        `;
    });
    
    grid.innerHTML = html;
}

function renderCategories() {
    const container = document.getElementById('categories-container');
    let html = '';

    categoriesData.forEach((cat) => {
        const isActive = cat === currentCategory;
        
        // Active = Dark/Black, Inactive = White pill with border
        const baseClass = "px-6 py-2.5 rounded-full text-sm font-semibold transition-all cursor-pointer border select-none";
        const activeClass = "bg-slate-800 dark:bg-brand-600 text-white border-slate-800 dark:border-brand-600 shadow-lg shadow-slate-200 dark:shadow-none";
        const inactiveClass = "bg-white dark:bg-slate-800 text-slate-600 dark:text-slate-300 border-slate-200 dark:border-slate-700 hover:border-slate-300 hover:text-slate-900 dark:hover:text-white";

        // We use onclick="setCategory('${cat}')" to trigger the update
        html += `<button onclick="setCategory('${cat}')" class="${baseClass} ${isActive ? activeClass : inactiveClass}">${cat}</button>`;
    });

    container.innerHTML = html;
}

// 3. ACTION HANDLER
function setCategory(category) {
    currentCategory = category;
    renderCategories(); // Re-render tabs to show active state
    renderTools();      // Re-render grid with new filter
}

// Initialize when the DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    renderTools();
    renderCategories();
});