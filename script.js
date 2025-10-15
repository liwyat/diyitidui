document.addEventListener('DOMContentLoaded', () = {
  const uploadForm = document.getElementById('uploadForm');
  const resultDiv = document.getElementById('result');
  const downloadBtn = document.getElementById('downloadBtn');
  let processedData = null;  存储处理后的数据，用于导出

   表单提交处理
  uploadForm.addEventListener('submit', (e) = {
    e.preventDefault();
    const fileInput = document.getElementById('file');
    const file = fileInput.files[0];
    
    if (!file) {
      alert('请选择Excel文件');
      return;
    }

     读取目标值设置
    const targets = {
      satisfaction parseFloat(document.getElementById('satisfactionTarget').value),
      resolution parseFloat(document.getElementById('resolutionTarget').value),
      sampleEval parseInt(document.getElementById('sampleEvalTarget').value),
      sampleRes parseInt(document.getElementById('sampleResTarget').value)
    };

     读取Excel文件
    const reader = new FileReader();
    reader.onload = (e) = {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type 'array' });
        const firstSheet = workbook.SheetNames[0];
        const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet]);
        
         处理数据
        processedData = processData(jsonData, targets);
         展示结果
        renderResults(processedData);
        resultDiv.classList.remove('hidden');
      } catch (error) {
        alert('文件解析错误：' + error.message);
      }
    };
    reader.readAsArrayBuffer(file);
  });

   下载Excel
  downloadBtn.addEventListener('click', () = {
    if (!processedData) return;
    
     合并所有数据（保留原始字段 + 新增字段）
    const allData = [
      ...processedData.oldElite,
      ...processedData.oldNonElite,
      ...processedData.newElite,
      ...processedData.newNonRes
    ];

     创建工作簿并导出
    const worksheet = XLSX.utils.json_to_sheet(allData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '处理结果');
    
    const excelBuffer = XLSX.write(workbook, { bookType 'xlsx', type 'array' });
    const blob = new Blob([excelBuffer], { type 'applicationoctet-stream' });
    saveAs(blob, '员工数据处理结果.xlsx');
  });

   数据处理核心函数
  function processData(rawData, targets) {
     区分新老员工
    const oldEmployees = rawData.filter(row = row['新老员工'] === '老员工');
    const newEmployees = rawData.filter(row = row['新老员工'] === '新员工');

     处理老员工
    const oldProcessed = oldEmployees.map(row = {
       判断是否达标
      const isElite = 
        row['满意度_剔除R'] = targets.satisfaction &&
        row['解决率_剔除R'] = targets.resolution &&
        row['评价量_剔除R'] = targets.sampleEval &&
        row['解决率总评价量_剔除R'] = targets.sampleRes;

       计算未达标字段
      const unmet = [];
      if (row['满意度_剔除R']  targets.satisfaction) {
        unmet.push(`满意度 ${row['满意度_剔除R']} (差距 ${(targets.satisfaction - row['满意度_剔除R']).toFixed(1)})`);
      }
      if (row['解决率_剔除R']  targets.resolution) {
        unmet.push(`解决率 ${row['解决率_剔除R']} (差距 ${(targets.resolution - row['解决率_剔除R']).toFixed(1)})`);
      }
      if (row['评价量_剔除R']  targets.sampleEval) {
        unmet.push(`评价量 ${row['评价量_剔除R']} (差距 ${targets.sampleEval - row['评价量_剔除R']})`);
      }
      if (row['解决率总评价量_剔除R']  targets.sampleRes) {
        unmet.push(`解决率评价量 ${row['解决率总评价量_剔除R']} (差距 ${targets.sampleRes - row['解决率总评价量_剔除R']})`);
      }

      return { ...row, isElite, 未达标字段 unmet.join('; ') };
    });

     老员工分组
    const oldElite = oldProcessed.filter(row = row.isElite);
    const oldNonElite = oldProcessed.filter(row = !row.isElite);

     处理新员工
    const totalEvaluation = newEmployees.reduce((sum, row) = sum + (row['评价量_剔除R']  0), 0);
    const newProcessed = newEmployees.map(row = {
       计算满意度影响值
      const impact = totalEvaluation  0 
         ((row['满意度_剔除R']  0) - targets.satisfaction)  ((row['评价量_剔除R']  0)  totalEvaluation)
         0;
      
       解决率是否达标
      const res达标 = (row['解决率_剔除R']  0) = 70;
      const res差距 = res达标  ''  `差距 ${(70 - row['解决率_剔除R']).toFixed(1)}`;

      return { ...row, 满意度影响值 impact.toFixed(4), res达标, res差距 };
    });

     新员工分组（影响值前100）
    const newElite = [...newProcessed]
      .sort((a, b) = b.满意度影响值 - a.满意度影响值)
      .slice(0, 100);
    const newNonRes = newProcessed.filter(row = !row.res达标);

    return { oldElite, oldNonElite, newElite, newNonRes };
  }

   渲染结果表格
  function renderResults(data) {
     渲染老员工第一梯队
    renderTable('oldElite', data.oldElite, ['姓名', '满意度_剔除R', '解决率_剔除R', '评价量_剔除R', '解决率总评价量_剔除R']);
    
     渲染老员工未达标
    renderTable('oldNonElite', data.oldNonElite, [
      '姓名', '满意度_剔除R', '解决率_剔除R', '评价量_剔除R', '解决率总评价量_剔除R', '未达标字段'
    ], (row, field) = field === '未达标字段'  'highlight-error'  '');
    
     渲染新员工第一梯队
    renderTable('newElite', data.newElite, [
      '姓名', '满意度_剔除R', '评价量_剔除R', '满意度影响值', '解决率_剔除R', 'res差距'
    ]);
    
     渲染新员工解决率未达标
    renderTable('newNonRes', data.newNonRes, [
      '姓名', '解决率_剔除R', 'res差距', '满意度_剔除R', '评价量_剔除R', '满意度影响值'
    ], (row, field) = field === 'res差距'  'highlight-error'  '');
  }

   通用表格渲染函数
  function renderTable(prefix, data, fields, getClass = () = '') {
    const headerEl = document.getElementById(`${prefix}Header`);
    const bodyEl = document.getElementById(`${prefix}Body`);
    
     清空表格
    headerEl.innerHTML = '';
    bodyEl.innerHTML = '';

     渲染表头
    const headerRow = document.createElement('tr');
    fields.forEach(field = {
      const th = document.createElement('th');
      th.className = 'px-4 py-3 text-left border-b';
      th.textContent = field.replace('res差距', '解决率差距');  替换显示文本
      headerRow.appendChild(th);
    });
    headerEl.appendChild(headerRow);

     渲染表体
    if (data.length === 0) {
      const emptyRow = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = fields.length;
      td.className = 'px-4 py-3 text-center text-gray-500 border-b';
      td.textContent = '无数据';
      emptyRow.appendChild(td);
      bodyEl.appendChild(emptyRow);
      return;
    }

    data.forEach(row = {
      const tr = document.createElement('tr');
      fields.forEach(field = {
        const td = document.createElement('td');
        td.className = `px-4 py-3 border-b ${getClass(row, field)}`;
        td.textContent = row[field]  '-';
        tr.appendChild(td);
      });
      bodyEl.appendChild(tr);
    });
  }
});