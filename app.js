class ExerciseSystem {
  constructor() {
    this.questions = [];
    this.currentQuestion = null;
    this.answered = 0;
    this.correct = 0;
    this.skipped = 0;  // 已跳过计数
    this.initEventListeners();
  }

  initEventListeners() {
    document.getElementById('excelFile').addEventListener('change', e => this.handleFile(e));
    document.getElementById('submitBtn').addEventListener('click', () => this.checkAnswer());
    document.getElementById('restartBtn').addEventListener('click', () => this.reset());
    document.getElementById('skipBtn').addEventListener('click', () => this.handleSkip());  // 跳过按钮监听
  }

  async handleFile(event) {
    try {
      const file = event.target.files[0];
      const formData = new FormData();
      formData.append('file', file);
      
      // 新增：设置请求超时（5秒）
      const controller = new AbortController();
      // 修改为10秒超时（10000毫秒）
      const timeoutId = setTimeout(() => controller.abort(), 10000);
      
      const response = await fetch('http://localhost:5000/process_excel', {
        method: 'POST',
        body: formData,
        signal: controller.signal // 关联超时信号
      });
      clearTimeout(timeoutId); // 清除未触发的超时
      
      if (!response.ok) {
        throw new Error('服务器处理失败');
      }
      
      const data = await response.json();
      this.questions = data.map(item => ({
        name: item.name,
        image: item.image
      }));
      
      this.nextQuestion();
    } catch (error) {
      if (error.name === 'AbortError') {
        alert('请求超时，请检查后端服务或网络');
      } else if (error.message.includes('Failed to fetch')) {
        alert('无法连接到服务器，请确保后端服务正在运行并监听5000端口');
      } else {
        alert('文件解析失败: ' + error.message);
      }
      console.error('文件上传错误:', error);
    }
  }

  readExcel(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const workbook = XLSX.read(e.target.result, {type: 'binary'});
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        resolve(XLSX.utils.sheet_to_json(sheet));
      };
      reader.onerror = () => reject(new Error('文件读取失败'));
      reader.readAsBinaryString(file);
    });
  }

  processData(data) {
    if (!data || !data.length) {
      throw new Error('Excel文件为空或格式不正确，请确保文件包含数据且格式正确');
    }
    
    return data.map(row => {
      if (!row['姓名']) throw new Error('缺少必要字段：姓名');
      const image = this.extractImage(row);
      if (!image) {
        console.warn('图片解析失败 - 原始数据:', JSON.stringify(row));
        console.log('可用列:', Object.keys(row).join(','));
        throw new Error(`无法解析图片数据，请检查图片列是否包含'图'字或图片格式是否正确`);
      }
      return {
        name: row['姓名'].trim(),
        image: image
      };
    }).filter(q => q.name && q.image);
  }

  extractImage(row) {
    try {
      if (!row.sheet || !row.sheet['!images']) {
        console.warn('Excel行数据缺少sheet或图片元数据:', JSON.stringify(row));
        return null;
      }
      
      const sheet = row.sheet;
      console.log('当前行图片元数据:', sheet['!images']);
      
      // 查找包含'图'或'image'的列
      const picColumns = Object.keys(row).filter(k => 
        k.includes('图') || k.toLowerCase().includes('image')
      );
      
      if(picColumns.length === 0) {
        console.warn('未找到图片列，可用列:', Object.keys(row).join(','));
        return null;
      }
      
      // 尝试所有可能的图片列
      for (const col of picColumns) {
        const picColumnIndex = Object.keys(row).indexOf(col);
        const address = XLSX.utils.encode_cell({r: row.__rowNum__, c: picColumnIndex});
        const image = sheet['!images']?.find(img => img.position?.origin === address);
        
        if (image?.base64) return image.base64;
      }
      
      // 兼容多种旧格式
      const fallbackImg = Object.values(row).find(v => {
        if (!v) return false;
        // 检查二进制图片数据
        if (v?.w && (v.t === 'j' || v.t === 'p' || v.t === 'g')) {
          return true;
        }
        // 检查base64图片数据
        if (typeof v === 'string' && 
            (v.startsWith('data:image') || 
             /[A-Za-z0-9+/=]{20,}/.test(v))) {
          return true;
        }
        return false;
      });
      
      if (fallbackImg?.w) {
        const typeMap = {j: 'jpeg', p: 'png', g: 'gif'};
        return `data:image/${typeMap[fallbackImg.t] || 'jpeg'};base64,${fallbackImg.v}`;
      }
      
      if (typeof fallbackImg === 'string') {
        return fallbackImg.startsWith('data:image') ? 
          fallbackImg : 
          `data:image/jpeg;base64,${fallbackImg}`;
      }
      
      console.warn('无法解析任何图片格式，原始数据:', JSON.stringify(row));
      return null;
    } catch (e) {
      console.error('图片解析错误:', e, '原始数据:', JSON.stringify(row));
      return null;
    }
  }

  nextQuestion() {
    if (!this.questions.length) {
      alert('有效题目为空，请检查文件格式');
      return;
    }
    this.currentQuestion = this.questions[Math.floor(Math.random() * this.questions.length)];
    const imgSrc = this.currentQuestion.image || 'placeholder.jpg';
    document.getElementById('preview-area').innerHTML = 
      `<img src="${imgSrc}" class="img-fluid" style="max-height: 300px">`;
    document.getElementById('answerInput').value = '';
  }

  checkAnswer() {
    const input = document.getElementById('answerInput').value.trim();
    if (!input) return;

    this.answered++;
    const isCorrect = input === this.currentQuestion.name;
    if (isCorrect) this.correct++;

    document.getElementById('preview-area').classList.add(isCorrect ? 'correct' : 'wrong');
    setTimeout(() => {
      document.getElementById('preview-area').classList.remove('correct', 'wrong');
      this.nextQuestion();
    }, 1000);

    this.updateStats();
  }

  updateStats() {
    document.getElementById('answeredCount').textContent = this.answered;
    document.getElementById('skippedCount').textContent = this.skipped;
    const accuracy = this.answered ? (this.correct / this.answered * 100).toFixed(1) : 0;
    document.getElementById('accuracyRate').textContent = `${accuracy}%`;
  }

  reset() {
    this.questions = [];
    this.currentQuestion = null;
    this.answered = 0;
    this.correct = 0;
    this.skipped = 0;
    document.getElementById('excelFile').value = '';
    document.getElementById('preview-area').innerHTML = '';
    this.updateStats();
  }

  handleSkip() {
    if (!this.currentQuestion) return;

    // 保持图片原有布局，在图片右侧叠加答案
    const previewArea = document.getElementById('preview-area');
    previewArea.innerHTML = `
        <div style="position: relative; display: inline-block;">
            <img src="${this.currentQuestion.image}" class="img-fluid" style="max-height: 300px">
            <div id="highlight-answer" style="position: absolute; 
                                          left: 105%; 
                                          top: 50%;
                                          transform: translateY(-50%);
                                          color: #28a745; 
                                          font-weight: bold; 
                                          font-size: 1.2rem;
                                          opacity: 1;
                                          transition: opacity 0.5s;
                                          white-space: nowrap;">
                正确答案：${this.currentQuestion.name}
            </div>
        </div>
    `;

    this.skipped++;
    this.updateStats();

    // 3 秒后淡出答案并切换题目
    setTimeout(() => {
        const answerDiv = document.getElementById('highlight-answer');
        if (answerDiv) {
            answerDiv.style.opacity = '0';
            setTimeout(() => this.nextQuestion(), 500);
        } else {
            this.nextQuestion();
        }
    }, 3000);
  }
}

// 仅保留一个初始化
new ExerciseSystem();