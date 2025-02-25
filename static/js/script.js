document.addEventListener('DOMContentLoaded', function() {
    // File input validation
    const fileInput = document.getElementById('file');
    fileInput.addEventListener('change', function() {
        const file = this.files[0];
        if (file) {
            const extension = file.name.split('.').pop().toLowerCase();
            if (extension !== 'docx') {
                alert('Please select a valid Word document (.docx)');
                this.value = '';
            }
        }
    });

    // Form submission handling
    const form = document.getElementById('uploadForm');
    form.addEventListener('submit', function(e) {
        const fileInput = document.getElementById('file');
        if (!fileInput.files.length) {
            e.preventDefault();
            alert('Please select a file to upload');
        }
    });
});

// Function to handle text file generation
function generateTextFiles() {
    const fileInput = document.getElementById('file');
    if (!fileInput.files.length) {
        alert('Please select a file first');
        return;
    }

    // Show loading indicator
    const btn = document.querySelector('button[onclick="generateTextFiles()"]');
    const originalText = btn.innerHTML;
    btn.innerHTML = '<i data-feather="loader" class="loader"></i> Generating...';
    btn.disabled = true;

    // Create and submit form
    const form = document.createElement('form');
    form.method = 'POST';
    form.action = '/generate-text-files';
    form.style.display = 'none';

    const fileInputClone = fileInput.cloneNode(true);
    form.appendChild(fileInputClone);

    document.body.appendChild(form);
    form.submit();

    // Reset button after a short delay
    setTimeout(() => {
        btn.innerHTML = originalText;
        btn.disabled = false;
        feather.replace();
    }, 2000);
}

// Add loading animation styles
const style = document.createElement('style');
style.textContent = `
    .loader {
        animation: spin 1s linear infinite;
    }
    @keyframes spin {
        100% { transform: rotate(360deg); }
    }
`;
document.head.appendChild(style);