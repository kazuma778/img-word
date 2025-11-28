class FileManager {
    constructor(fileInputId, fileListId, fileDisplayClass, options = {}) {
        this.fileInput = document.getElementById(fileInputId);
        this.fileList = document.getElementById(fileListId);
        this.fileDisplay = document.querySelector('.' + fileDisplayClass);
        this.files = []; // Array to store File objects
        this.options = options; // e.g., { onUpdate: callback }

        if (!this.fileInput || !this.fileList || !this.fileDisplay) {
            console.error("FileManager: Required elements not found.");
            return;
        }

        this.init();
    }

    init() {
        // Handle file input change
        this.fileInput.addEventListener('change', (e) => {
            this.addFiles(e.target.files);
            // Do NOT clear value here, as it wipes out the files we just set in updateInput()
            // this.fileInput.value = ''; 
        });

        // Handle drag and drop
        this.fileDisplay.addEventListener('dragover', (e) => {
            e.preventDefault();
            this.fileDisplay.classList.add('dragover');
        });

        this.fileDisplay.addEventListener('dragleave', () => {
            this.fileDisplay.classList.remove('dragover');
        });

        this.fileDisplay.addEventListener('drop', (e) => {
            e.preventDefault();
            this.fileDisplay.classList.remove('dragover');
            this.addFiles(e.dataTransfer.files);
        });
    }

    addFiles(newFiles) {
        if (!newFiles || newFiles.length === 0) return;

        Array.from(newFiles).forEach(file => {
            // Optional: Check for duplicates? For now, we allow duplicates as user might want to merge same file twice.
            // But usually for processing it's better to avoid. Let's allow it for flexibility.
            this.files.push(file);
        });

        this.render();
        this.updateInput();
    }

    removeFile(index) {
        this.files.splice(index, 1);
        this.render();
        this.updateInput();
    }

    moveUp(index) {
        if (index > 0) {
            [this.files[index], this.files[index - 1]] = [this.files[index - 1], this.files[index]];
            this.render();
            this.updateInput();
        }
    }

    moveDown(index) {
        if (index < this.files.length - 1) {
            [this.files[index], this.files[index + 1]] = [this.files[index + 1], this.files[index]];
            this.render();
            this.updateInput();
        }
    }

    updateInput() {
        // Create a new DataTransfer object to sync with the input
        const dataTransfer = new DataTransfer();
        this.files.forEach(file => dataTransfer.items.add(file));
        this.fileInput.files = dataTransfer.files;

        if (this.options.onUpdate) {
            this.options.onUpdate(this.files);
        }
    }

    getFiles() {
        return this.files;
    }

    render() {
        if (this.files.length === 0) {
            this.fileList.style.display = 'none';
            this.fileList.innerHTML = '';
            return;
        }

        this.fileList.style.display = 'block';
        this.fileList.innerHTML = '';

        this.files.forEach((file, index) => {
            const item = document.createElement('div');
            item.className = 'file-manager-item';

            const nameSpan = document.createElement('span');
            nameSpan.className = 'file-name';
            nameSpan.textContent = `📄 ${index + 1}. ${file.name} (${this.formatSize(file.size)})`;

            const actionsDiv = document.createElement('div');
            actionsDiv.className = 'file-actions';

            // Up Button
            const upBtn = document.createElement('button');
            upBtn.type = 'button';
            upBtn.className = 'btn-icon btn-sm';
            upBtn.innerHTML = '⬆️';
            upBtn.title = 'Geser ke atas';
            upBtn.disabled = index === 0;
            upBtn.onclick = () => this.moveUp(index);

            // Down Button
            const downBtn = document.createElement('button');
            downBtn.type = 'button';
            downBtn.className = 'btn-icon btn-sm';
            downBtn.innerHTML = '⬇️';
            downBtn.title = 'Geser ke bawah';
            downBtn.disabled = index === this.files.length - 1;
            downBtn.onclick = () => this.moveDown(index);

            // Remove Button
            const removeBtn = document.createElement('button');
            removeBtn.type = 'button';
            removeBtn.className = 'btn-icon btn-sm btn-danger';
            removeBtn.innerHTML = '❌';
            removeBtn.title = 'Hapus file';
            removeBtn.onclick = () => this.removeFile(index);

            actionsDiv.appendChild(upBtn);
            actionsDiv.appendChild(downBtn);
            actionsDiv.appendChild(removeBtn);

            item.appendChild(nameSpan);
            item.appendChild(actionsDiv);
            this.fileList.appendChild(item);
        });
    }

    formatSize(bytes) {
        if (bytes === 0) return '0 B';
        const k = 1024;
        const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }
}
