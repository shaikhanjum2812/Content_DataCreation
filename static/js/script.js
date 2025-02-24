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

    const formData = new FormData();
    formData.append('file', fileInput.files[0]);

    const form = document.createElement('form');
    form.method = 'POST';
    form.action = '/generate-text-files';
    form.style.display = 'none';

    const fileInputClone = fileInput.cloneNode(true);
    form.appendChild(fileInputClone);

    document.body.appendChild(form);
    form.submit();
    document.body.removeChild(form);
}
