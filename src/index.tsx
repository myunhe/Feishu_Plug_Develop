import React, { useEffect, useState } from 'react'
import ReactDOM from 'react-dom/client'
import { bitable, CurrencyCode, FieldType, ICurrencyField, ICurrencyFieldMeta, IFieldMeta} from '@lark-base-open/js-sdk';
import { Alert, AlertProps, Button, Select, Tabs, List, Tag, Space, Input, Popconfirm, message, Modal } from 'antd';
import { SettingOutlined, PlusOutlined, DeleteOutlined } from '@ant-design/icons';
import { CURRENCY } from './const';
import { getExchangeRate } from './exchange-api';

// API配置 - 全局变量，便于统一修改
const API_BASE_URL = 'https://salvational-unvisible-ligia.ngrok-free.dev';

// 全局日志控制
const DEBUG = false; // 设置为 false 可关闭所有调试日志

const debuglog = (...args: any[]) => {
  if (DEBUG) {
    console.log(...args);
  }
};

ReactDOM.createRoot(document.getElementById('root') as HTMLElement).render(
  <React.StrictMode>
    <LoadApp/>
  </React.StrictMode>
)

function LoadApp() {
  const [info, setInfo] = useState('get table name, please waiting ....');
  const [alertType, setAlertType] = useState<AlertProps['type']>('info');
  const [currencyFieldMetaList, setMetaList] = useState<ICurrencyFieldMeta[]>([])
  const [selectFieldId, setSelectFieldId] = useState<string>();
  const [currency, setCurrency] = useState<CurrencyCode>();
  const [selectedProject, setSelectedProject] = useState<string>(() => {
    // 从localStorage加载上次选择的项目
    const savedProject = localStorage.getItem('feishu_selected_project');
    return savedProject || '';
  });
  const [allFieldMetaList, setAllFieldMetaList] = useState<IFieldMeta[]>([]);
  const [selectedValueField, setSelectedValueField] = useState<string>(() => {
    // 从localStorage加载上次选择的实际值列
    const savedValueField = localStorage.getItem('feishu_selected_value_field');
    return savedValueField || '';
  });
  const [errorMessages, setErrorMessages] = useState<string[]>([]);
  const [conversionStatus, setConversionStatus] = useState<'idle' | 'converting' | 'success' | 'failed' | 'warning'>('idle');
  const [statusMessage, setStatusMessage] = useState<string>('');
  
  const [projectOptions, setProjectOptions] = useState<{ label: string; value: string }[]>([]);
  const [newProjectName, setNewProjectName] = useState<string>('');
  const [showProjectManager, setShowProjectManager] = useState<boolean>(false);
  const [tempNewProjectName, setTempNewProjectName] = useState<string>('');

  // 从后端获取项目列表的函数
  const fetchProjectListFromBackend = async () => {
    try {
      // 首先检查localStorage中是否有保存的项目列表
      const savedProjects = localStorage.getItem('feishu_project_options');
      if (savedProjects) {
        debuglog('从localStorage加载项目列表');
        const projects = JSON.parse(savedProjects);
        setProjectOptions(projects);
        return;
      }
      
      debuglog('localStorage中没有项目列表，开始从后端获取...');
      // 添加ngrok特定的请求头
      const response = await fetch(`${API_BASE_URL}/api/get-project-list`, {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json',
          'ngrok-skip-browser-warning': 'true', // 跳过ngrok浏览器警告
        },
        mode: 'cors', // 明确指定CORS模式
      });
      
      debuglog("项目列表响应状态:", response.status, response.statusText);
      debuglog("响应是否成功:", response.ok);
      debuglog("响应头:", Object.fromEntries(response.headers.entries()));
      
      if (!response.ok) {
        console.error('HTTP错误:', response.status, response.statusText);
        
        // 如果是重定向错误，尝试直接访问本地地址
        if (response.status >= 300 && response.status < 400) {
          debuglog('检测到重定向，尝试访问本地地址...');
          // await fetchProjectListFromLocal();
        }
        return;
      }
      
      const responseText = await response.text();
      debuglog("原始响应文本:", responseText);
      
      // 清理响应文本，移除可能的HTML或非JSON内容
      const cleanedText = cleanResponseText(responseText);
      debuglog("清理后的响应文本:", cleanedText);
      
      try {
        const data = JSON.parse(cleanedText);
        debuglog("解析后的项目列表数据:", data);
        
        if (data.success) {
          if (data.projects && Array.isArray(data.projects)) {
            debuglog("后端返回项目数量:", data.projects.length);
            
            // 直接使用后端返回的项目格式，不需要转换
            const backendProjects = data.projects.map((project: any) => ({
              label: project.label || project.name || `项目${project.id}`,
              value: project.value || project.id || `project_${Date.now()}`
            }));
            
            debuglog("处理后的项目列表:", backendProjects);
            
            // 直接使用后端项目列表，不合并默认项目
            setProjectOptions(backendProjects);
            
            // 保存到localStorage
            localStorage.setItem('feishu_project_options', JSON.stringify(backendProjects));
            
            debuglog('项目列表获取并保存成功');
          } else {
            console.error('项目数据格式错误: projects字段不存在或不是数组', data);
          }
        } else {
          console.error('后端返回失败:', data);
        }
      } catch (jsonError) {
        console.error('JSON解析错误:', jsonError);
        console.error('原始响应内容:', responseText);
        console.error('清理后的内容:', cleanedText);
      }
    } catch (error) {
      console.error('获取项目列表失败:', error);
    }
  };

  // 清理响应文本的函数
  const cleanResponseText = (text: string): string => {
    // 移除HTML标签
    let cleaned = text.replace(/<[^>]*>/g, '');
    
    // 查找JSON对象开始位置
    const jsonStart = cleaned.indexOf('{');
    const jsonEnd = cleaned.lastIndexOf('}');
    
    if (jsonStart !== -1 && jsonEnd !== -1 && jsonEnd > jsonStart) {
      cleaned = cleaned.substring(jsonStart, jsonEnd + 1);
    }
    
    // 移除多余的空格和换行
    cleaned = cleaned.trim();
    
    return cleaned;
  };

  // 从本地地址获取项目列表的备用函数
  const fetchProjectListFromLocal = async () => {
    try {
      debuglog('尝试从本地地址获取项目列表...');
      const response = await fetch('http://localhost:5005/api/get-project-list', {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json',
        },
      });
      
      if (response.ok) {
        const data = await response.json();
        if (data.projects && Array.isArray(data.projects)) {
          const backendProjects = data.projects.map((project: any) => ({
            label: project.label || project.name || `项目${project.id}`,
            value: project.value || project.id || `project_${Date.now()}`
          }));
          
          // 直接使用后端项目列表，不合并默认项目
          setProjectOptions(backendProjects);
          localStorage.setItem('feishu_project_options', JSON.stringify(backendProjects));
          debuglog('从本地地址获取项目列表成功');
        }
      }
    } catch (error) {
      console.error('从本地地址获取项目列表失败:', error);
    }
  };

  useEffect(() => {
    const fn = async () => {
      const table = await bitable.base.getActiveTable();
      const tableName = await table.getName();
      setInfo(`The table Name is ${tableName}`);
      setAlertType('success');
      const fieldMetaList = await table.getFieldMetaListByType<ICurrencyFieldMeta>(FieldType.Currency);
      setMetaList(fieldMetaList);
      const allFields = await table.getFieldMetaList();
      debuglog("所有字段:", allFields);
      // setAllFieldMetaList(allFields);

      const selection = await bitable.base.getSelection();
      const viewId = selection.viewId;
      const view = await table.getViewById(viewId as string);
      const visibleFieldIdList = await view.getVisibleFieldIdList();
      const orderedFields = visibleFieldIdList.map(fieldId => {
        const field = allFields.find(f => f.id === fieldId);
        if (field) {
          return field;
        }
        return null;
      }).filter(field => field !== null) as IFieldMeta[];
      setAllFieldMetaList(orderedFields);

      // 从后端获取项目列表
      await fetchProjectListFromBackend();
    };
    fn();
  }, []);

  const formatFieldMetaList = (metaList: ICurrencyFieldMeta[]) => {
    return metaList.map(meta => ({ label: meta.name, value: meta.id }));
  };

  const formatAllFieldMetaList = (metaList: IFieldMeta[]) => {
    return metaList.map(meta => ({ label: meta.name, value: meta.id }));
  };

  const transform = async () => {
    if (!selectFieldId || !currency) return;
    const table = await bitable.base.getActiveTable();
    const currencyField = await table.getField<ICurrencyField>(selectFieldId);
    const currentCurrency = await currencyField.getCurrencyCode();
    await currencyField.setCurrencyCode(currency);
    const ratio = await getExchangeRate(currentCurrency, currency);
    if (!ratio) return;
    const recordIdList = await table.getRecordIdList();
    for (const recordId of recordIdList) {
      const currentVal = await currencyField.getValue(recordId);
      await currencyField.setValue(recordId, currentVal * ratio);
    }
  }

  const startMessageConversion = async () => {
    if (!selectedProject || !selectedValueField) {
      setErrorMessages(['请选择项目和实际值列']);
      setConversionStatus('failed');
      setStatusMessage('转换失败：参数不完整');
      return;
    }

    setConversionStatus('converting');
    setStatusMessage('正在获取当前视图数据...');
    setErrorMessages([]);

    try {
      // 获取当前表格
      const selection = await bitable.base.getSelection();
      const tableId = selection.tableId;
      const viewId = selection.viewId;

      const table = await bitable.base.getTableById(tableId as string);
      const view = await table.getViewById(viewId as string);
      const view_meta = await table.getViewMetaById(viewId as string);
      const tableName = await table.getName();

      debuglog("---------------->view name:", view.name)
      debuglog("---------------->view meta:", view_meta.name)
      debuglog("---------------->view meta property:", view_meta.property)
      const visibleFieldIdList = await view.getVisibleFieldIdList();
      debuglog("---------------->11view meta property:", visibleFieldIdList)
      
      const recordList_view = await view.getVisibleRecordIdList();
      debuglog('recordList_view:', recordList_view);
      
      // 获取所有字段的元数据
      const fieldMetaList = await table.getFieldMetaList();
      debuglog("---------------->fieldMetaList:", fieldMetaList)
      
      // 检查表头是否包含所有必需字段
      const requiredHeaders = ["DID", "名称", "二级名称", "英文名称", "类型", "长度", "读SID", "读Session", "写SID", "写Session", "数据格式"];
      const currentHeaders = fieldMetaList.map(field => field.name);
      
      const missingHeaders = requiredHeaders.filter(header => !currentHeaders.includes(header));
      
      if (missingHeaders.length > 0) {
        setErrorMessages([`当前表格格式错误，缺少以下必需字段：${missingHeaders.join(", ")}`]);
        setConversionStatus('failed');
        setStatusMessage('转换失败：表格格式不正确');
        return;
      }

      // 获取所有记录的数据并转换为二维数组格式
      const feishuDataArray = [];
      
      // 添加表头（字段名）
      const headers = fieldMetaList.map(field => field.name);
      feishuDataArray.push(headers);
      
      // 获取所有记录的数据
      for (const recordId of recordList_view) {
        const record = await table.getRecordById(recordId as string);
        const row = [];
        
        // 按照字段顺序获取每个字段的值
        for (const field of fieldMetaList) {
          const fieldValue = record.fields[field.id] || '';
          let textValue = '';
          
          // 根据字段类型提取文本值
          if (fieldValue) {
            if (Array.isArray(fieldValue)) {
              // 处理数组类型的字段
              const extractedValues = [];
              for (const item of fieldValue) {
                if (item && typeof item === 'object' && item !== null) {
                  // 格式1: {id: 'XX', text: '否'} - 提取text
                  if ('text' in item && item.text !== undefined) {
                    extractedValues.push(String(item.text));
                  }
                  // 格式2: {id: 'XX', name: 'XX'} - 提取name
                  else if ('name' in item && item.name !== undefined) {
                    extractedValues.push(String(item.name));
                  }
                  // 其他对象格式
                  else {
                    extractedValues.push(JSON.stringify(item));
                  }
                } else {
                  // 直接值
                  extractedValues.push(String(item));
                }
              }
              textValue = extractedValues.join(', ');
            } else if (typeof fieldValue === 'object' && fieldValue !== null) {
              // 处理单个对象类型的字段
              // 格式1: {id: 'XX', text: '否'} - 提取text
              if ('text' in fieldValue && fieldValue.text !== undefined) {
                textValue = String(fieldValue.text);
              }
              // 格式2: {id: 'XX', name: 'XX'} - 提取name
              else if ('name' in fieldValue && fieldValue.name !== undefined) {
                textValue = String(fieldValue.name);
              }
              // 其他对象格式
              else {
                textValue = JSON.stringify(fieldValue);
              }
            } else {
              // 直接转换为字符串
              textValue = String(fieldValue);
            }
          }
          
          row.push(textValue);
        }
        
        feishuDataArray.push(row);
      }
      
      setStatusMessage(`成功获取 ${feishuDataArray.length - 1} 条记录数据，正在准备转换...`);
      
      // 显示获取到的数据信息（用于调试）
      debuglog('获取到的飞书数据（二维数组格式）:', feishuDataArray);
      debuglog('字段元数据:', fieldMetaList);
      
      // 调用Python后端API进行数据转换
      setStatusMessage('正在进行数据转换...');
      
      // 获取选中的实际值列的字段名
      const valueFieldMeta = allFieldMetaList.find(field => field.id === selectedValueField);
      const valueFieldName = valueFieldMeta ? valueFieldMeta.name : '实际值';
      
      // 准备配置项目信息
      const configProject = {
        sheet_name: '转换结果',
        value_column: valueFieldName,
        did: 'DID',
        did_name: '名称',
        signal_chinese: '信号中文名',
        signal_english: '信号英文名',
        type: '类型',
        length: '长度',
        read_sid: '读SID',
        read_session: '读Session',
        write_sid: '写SID',
        write_session: '写Session',
        data_format: '数据格式'
      };
      
      // 调用后端API
      const response = await fetch(`${API_BASE_URL}/api/convert-feishu-data`, 
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          feishu_data: JSON.stringify(feishuDataArray),
          config_project: configProject,
          project_name: selectedProject,
          view_name: view_meta.name,
          table_name: tableName
        })
      });
      
      const result = await response.json();
      if (result.success) {
        // 处理返回的Excel数据（十六进制格式）
        if (result.excel_data) {
          // 将十六进制字符串转换回二进制数据
          const hexString: string = result.excel_data;
          const hexMatches = hexString.match(/[\da-f]{2}/gi);
          
          if (hexMatches) {
            const byteArray = new Uint8Array(hexMatches.map(h => parseInt(h, 16)));
            
            // 创建Blob对象
            const blob = new Blob([byteArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            
            // 创建下载链接
            const downloadLink = document.createElement('a');
            downloadLink.href = URL.createObjectURL(blob);
            
            // 生成文件名
            const timestamp = new Date().toLocaleString('zh-CN').replace(/[\/\s:]/g, '_');
            // 根据selectedProject的值获取项目名称（显示在下拉框中的文本）
            const selectedProjectOption = projectOptions.find(project => project.value === selectedProject);
            const projectNameForFile = selectedProjectOption ? selectedProjectOption.label : selectedProject;
            downloadLink.download = `${projectNameForFile}_${view_meta.name}_${timestamp}_${tableName}.xlsx`;
            debuglog("---------------->project name:", projectNameForFile)
            debuglog("---------------->view name:", view.getName())
            debuglog("---------------->table name:", tableName)
            
            // 触发下载
            document.body.appendChild(downloadLink);
            downloadLink.click();
            document.body.removeChild(downloadLink);
            
            // 设置状态为转换成功
            setConversionStatus('success');
            setStatusMessage('转换成功，文件已下载');
          } else {
            throw new Error('十六进制数据格式错误');
          }
        } else {
          throw new Error('未收到Excel数据');
        }
      } else {
        throw new Error(result.error || '转换失败');
      }
      
    } catch (error) {
      setErrorMessages([`获取数据过程中发生错误: ${error}`]);
      setConversionStatus('failed');
      setStatusMessage('获取数据失败');
    }
  }

  const addProject = (projectName: string) => {
    if (!projectName.trim()) {
      message.warning('请输入项目名称');
      return;
    }
    
    const existingProject = projectOptions.find(p => p.label === projectName.trim());
    if (existingProject) {
      setSelectedProject(existingProject.value);
      message.info('已选择现有项目');
      return;
    }
    
    const newProject = {
      label: projectName.trim(),
      value: `project_${Date.now()}`
    };
    
    const updatedProjects = [...projectOptions, newProject];
    setProjectOptions(updatedProjects);
    localStorage.setItem('feishu_project_options', JSON.stringify(updatedProjects));
    setSelectedProject(newProject.value);
    message.success('项目添加成功');
  };

  const deleteProject = (projectValue: string) => {
    const updatedProjects = projectOptions.filter(p => p.value !== projectValue);
    setProjectOptions(updatedProjects);
    localStorage.setItem('feishu_project_options', JSON.stringify(updatedProjects));
    
    if (selectedProject === projectValue) {
      setSelectedProject('');
      localStorage.removeItem('feishu_selected_project');
    }
    
    message.success('项目删除成功');
  };

  // 处理项目选择变化并保存到localStorage
  const handleProjectChange = (value: string) => {
    setSelectedProject(value);
    if (value) {
      localStorage.setItem('feishu_selected_project', value);
    } else {
      localStorage.removeItem('feishu_selected_project');
    }
  };

  // 处理实际值列选择变化并保存到localStorage
  const handleValueFieldChange = (value: string) => {
    setSelectedValueField(value);
    if (value) {
      localStorage.setItem('feishu_selected_value_field', value);
    } else {
      localStorage.removeItem('feishu_selected_value_field');
    }
  };

  const tabItems = [
    {
      key: 'message',
      label: '报文转换',
      children: (
        <div style={{ display: 'flex', flexDirection: 'column', height: '400px' }}>
          <div style={{ padding: '20px', flex: 1, overflow: 'auto' }}>
            <div style={{ marginBottom: 16 }}>
              <div style={{ marginBottom: 8 }}>选择项目</div>
              <div style={{ display: 'flex', gap: '10px', alignItems: 'flex-end' }}>
                <Select
                  style={{ width: 200 }}
                  value={selectedProject}
                  onChange={handleProjectChange}
                  options={projectOptions}
                  placeholder="请选择项目"
                  showSearch
                  allowClear
                  filterOption={(input, option) => 
                    (option?.label ?? '').toLowerCase().includes(input.toLowerCase())
                  }
                />
                <Button 
                  type="text" 
                  icon={<SettingOutlined />}
                  onClick={() => setShowProjectManager(true)}
                  style={{ color: '#666' }}
                />
              </div>
            </div>
            <div style={{ marginBottom: 16 }}>
              <div style={{ marginBottom: 8 }}>实际值列</div>
              <Select 
                style={{ width: 200 }} 
                value={selectedValueField}
                onChange={handleValueFieldChange} 
                options={formatAllFieldMetaList(allFieldMetaList)} 
                placeholder="请选择实际值列"
              />
            </div>
            <div style={{ marginBottom: 16 }}>
              <Button 
                type="primary" 
                style={{ backgroundColor: '#1890ff', borderColor: '#1890ff', color: 'white' }}
                onClick={startMessageConversion}
                loading={conversionStatus === 'converting'}
              >
                开始转换
              </Button>
            </div>
            
            {errorMessages.length > 0 && (
              <div style={{ marginTop: 20 }}>
                <div style={{ marginBottom: 8, fontWeight: 'bold' }}>错误信息列表：</div>
                <List
                  size="small"
                  bordered
                  dataSource={errorMessages}
                  renderItem={(item, index) => (
                    <List.Item style={{ color: '#ff4d4f' }}>
                      {index + 1}. {item}
                    </List.Item>
                  )}
                />
              </div>
            )}
          </div>
          
          <div style={{ 
            borderTop: '1px solid #d9d9d9', 
            padding: '10px 20px', 
            backgroundColor: '#f5f5f5',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between'
          }}>
            <Space>
              <div>状态：</div>
              <Tag color={
                conversionStatus === 'converting' ? 'blue' :
                conversionStatus === 'success' ? 'green' :
                conversionStatus === 'failed' ? 'red' :
                conversionStatus === 'warning' ? 'orange' : 'default'
              }>
                {conversionStatus === 'idle' && '空闲'}
                {conversionStatus === 'converting' && '正在转换中'}
                {conversionStatus === 'success' && '转换成功'}
                {conversionStatus === 'failed' && '转换失败'}
                {conversionStatus === 'warning' && '转换警告'}
              </Tag>
            </Space>
            <div style={{ color: '#666' }}>{statusMessage}</div>
          </div>
          
          {/* 项目设置浮动窗口 */}
          <Modal
            title="项目管理"
            open={showProjectManager}
            onCancel={() => {
              setShowProjectManager(false);
              setTempNewProjectName('');
            }}
            footer={null}
            width={400}
            style={{ top: 20 }}
          >
            <div style={{ maxHeight: '300px', overflow: 'auto', marginBottom: '16px' }}>
              {projectOptions.length === 0 ? (
                <div style={{ textAlign: 'center', color: '#999', padding: '20px' }}>
                  暂无项目
                </div>
              ) : (
                <List
                  size="small"
                  dataSource={projectOptions}
                  renderItem={(project) => (
                    <List.Item
                      actions={[
                        <Popconfirm
                          title="确定删除这个项目吗？"
                          onConfirm={() => deleteProject(project.value)}
                          okText="确定"
                          cancelText="取消"
                        >
                          <Button type="text" size="small" icon={<DeleteOutlined />} />
                        </Popconfirm>
                      ]}
                    >
                      <div style={{ flex: 1 }}>{project.label}</div>
                    </List.Item>
                  )}
                />
              )}
            </div>
            
            <div style={{ borderTop: '1px solid #d9d9d9', paddingTop: '16px' }}>
              <Space.Compact style={{ width: '100%' }}>
                <Input
                  value={tempNewProjectName}
                  onChange={(e) => setTempNewProjectName(e.target.value)}
                  placeholder="输入新项目名称"
                  onPressEnter={() => {
                    addProject(tempNewProjectName);
                    setTempNewProjectName('');
                  }}
                />
                <Button 
                  type="primary" 
                  icon={<PlusOutlined />}
                  onClick={() => {
                    addProject(tempNewProjectName);
                    setTempNewProjectName('');
                  }}
                >
                  添加
                </Button>
              </Space.Compact>
            </div>
          </Modal>
        </div>
      ),
    },
    // {
    //   key: 'currency',
    //   label: '货币转换',
    //   children: (
    //     <div style={{ padding: '20px' }}>
    //       <div style={{ marginBottom: 16 }}>
    //         <div>Select Field</div>
    //         <Select style={{ width: 120 }} onSelect={setSelectFieldId} options={formatFieldMetaList(currencyFieldMetaList)}/>
    //       </div>
    //       <div style={{ marginBottom: 16 }}>
    //         <div>Select Currency</div>
    //         <Select options={CURRENCY} style={{ width: 120 }} onSelect={setCurrency}/>
    //       </div>
    //       <div>
    //         <Button style={{ marginLeft: 10 }} onClick={transform}>transform</Button>
    //       </div>
    //     </div>
    //   ),
    // },
  ];

  return <Tabs items={tabItems} tabBarStyle={{ display: 'none' }} />
}