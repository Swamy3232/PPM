import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import {
  DeleteOutlined,
  EditOutlined,
  PlusOutlined,
  SearchOutlined,
  DownloadOutlined,
  EyeOutlined,
} from '@ant-design/icons'
import {
  Button,
  Form,
  Input,
  Modal,
  Popconfirm,
  Space,
  Table,
  Typography,
  message,
  DatePicker,
  Select,
  Row,
  Col,
  Popover
} from 'antd'

import { ExcelRenderer } from 'react-excel-renderer'
import mammoth from 'mammoth'

import dayjs from 'dayjs'
import isSameOrAfter from 'dayjs/plugin/isSameOrAfter'
import isSameOrBefore from 'dayjs/plugin/isSameOrBefore'

import '../App.css'
import { API_BASE_URL } from '../config/api.js'
import { formatDate } from '../config/date.js'

dayjs.extend(isSameOrAfter)
dayjs.extend(isSameOrBefore)


const AcknowledgeProposalsTable = ({ fetchProposalsTrigger }) => {
  const [pendingProposals, setPendingProposals] = useState([])
  const [loading, setLoading] = useState(false)
  const [actionLoading, setActionLoading] = useState({})

  // Document modal state
  const [stageConfig, setStageConfig] = useState([])
  const [docsModalVisible, setDocsModalVisible] = useState(false)
  const [projectDocs, setProjectDocs] = useState([])
  const [docsLoading, setDocsLoading] = useState(false)
  const [viewDocumentUrl, setViewDocumentUrl] = useState(null)

  const [excelRendererData, setExcelRendererData] = useState(null)
  const [excelRendererLoading, setExcelRendererLoading] = useState(false)
  const [excelRendererError, setExcelRendererError] = useState(null)
  const [wordDocumentContent, setWordDocumentContent] = useState(null)
  const [wordDocumentLoading, setWordDocumentLoading] = useState(false)
  const [wordDocumentError, setWordDocumentError] = useState(null)
  const [currentUserRole, setCurrentUserRole] = useState('')

  const fetchPendingProposals = useCallback(async () => {
    setLoading(true)
    try {
      const response = await fetch(`${API_BASE_URL}/proposals/false`, {
        headers: { accept: 'application/json' },
      })
      if (!response.ok) throw new Error('Failed to fetch pending proposals')
      const data = await response.json()
      const normalized = Array.isArray(data)
        ? data.map((item) => ({ ...item, key: item.id }))
        : []

      // Fetch all documents to compute per-proposal document counts
      try {
        const docsRes = await fetch(`${API_BASE_URL}/documents/`, {
          headers: { accept: 'application/json' },
        })
        if (docsRes.ok) {
          const allDocs = await docsRes.json()
          const docsByProject = {}
            ; (Array.isArray(allDocs) ? allDocs : []).forEach((d) => {
              const pid = d.project_id
              if (pid != null) docsByProject[pid] = (docsByProject[pid] || 0) + 1
            })
          normalized.forEach((item) => {
            item._docCount = docsByProject[item.id] || 0
          })
        }
      } catch (docErr) {
        console.error('Failed to fetch document counts:', docErr)
      }

      setPendingProposals(normalized)
    } catch (error) {
      console.error(error)
      message.error('Unable to fetch pending proposals')
    } finally {
      setLoading(false)
    }
  }, [])

  useEffect(() => {
    try {
      const rawUser = window.localStorage.getItem('ppm_user')
      if (rawUser) {
        const parsedUser = JSON.parse(rawUser)
        if (parsedUser?.role) {
          setCurrentUserRole(parsedUser.role)
        }
      }
    } catch (error) {
      console.error('Failed to read user from localStorage', error)
    }
  }, [])

  const isGuest = currentUserRole?.toLowerCase().trim() === 'guest'

  useEffect(() => {
    fetchPendingProposals()
    fetchStageConfig()
  }, [fetchPendingProposals])

  const fetchStageConfig = useCallback(async () => {
    try {
      const res = await fetch(`${API_BASE_URL}/stages/`, {
        headers: { accept: 'application/json' },
      })
      if (!res.ok) throw new Error('Failed to fetch stage configuration')
      const data = await res.json()
      setStageConfig(Array.isArray(data) ? data.map((item) => ({ ...item, key: item.id })) : [])
    } catch (error) {
      console.error('Error fetching stage configuration:', error)
    }
  }, [])

  const fetchProjectDocuments = useCallback(async (projectId) => {
    setDocsLoading(true)
    try {
      const res = await fetch(`${API_BASE_URL}/documents/`, {
        headers: { accept: 'application/json' },
      })
      if (!res.ok) throw new Error('Failed to fetch documents')
      const data = await res.json()
      const docs = Array.isArray(data) ? data : []

      const enquiryStage = stageConfig.find(
        (s) => (s.name || '').toString().trim().toLowerCase() === 'enquiry',
      )
      const enquiryStageId = enquiryStage?.id

      const filtered = docs
        .filter((d) => d.project_id === projectId)
        .filter((d) => (enquiryStageId ? d.stage_id === enquiryStageId : true))

      const baseName = (enquiryStage?.name || 'Enquiry').toString().trim() || 'Enquiry'

      const sortedByDate = [...filtered].sort(
        (a, b) => new Date(a.created_at) - new Date(b.created_at),
      )

      const withVersions = sortedByDate.map((d, idx) => ({
        ...d,
        version: idx + 1,
        display_name: d.name || `${baseName} v${idx + 1}`,
      }))

      setProjectDocs(withVersions)
    } catch (err) {
      console.error('Error fetching project documents:', err)
      message.error(err.message || 'Unable to load documents')
      setProjectDocs([])
    } finally {
      setDocsLoading(false)
    }
  }, [stageConfig])

  const openDocsModal = useCallback(async (projectId) => {
    setDocsModalVisible(true)
    await fetchProjectDocuments(projectId)
  }, [fetchProjectDocuments])

  // const viewDocument = useCallback((doc) => {
  //   if (!doc?.url) {
  //     return message.error('Document URL is not available')
  //   }
  //   setViewDocumentUrl(doc.url)
  // }, [])

  const loadExcelWithRenderer = async (url) => {
    setExcelRendererLoading(true)
    setExcelRendererError(null)
    setExcelRendererData(null)
    try {
      const response = await fetch(url)
      if (!response.ok) throw new Error(`Failed to fetch Excel file: ${response.status}`)
      const blob = await response.blob()
      const file = new File([blob], 'excel.xlsx', {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      })
      ExcelRenderer(file, (err, resp) => {
        if (err) {
          setExcelRendererError(`Failed to parse Excel file: ${err.message || err}`)
          setExcelRendererLoading(false)
        } else {
          setExcelRendererData(resp)
          setExcelRendererLoading(false)
        }
      })
    } catch (error) {
      setExcelRendererError(`Error loading Excel file: ${error.message}`)
      setExcelRendererLoading(false)
    }
  }

  const loadWordDocument = async (url) => {
    setWordDocumentLoading(true)
    setWordDocumentError(null)
    setWordDocumentContent(null)
    try {
      const response = await fetch(url)
      if (!response.ok) throw new Error(`Failed to fetch Word document: ${response.status}`)
      const arrayBuffer = await response.arrayBuffer()
      const result = await mammoth.convertToHtml({ arrayBuffer })
      setWordDocumentContent(result.value)
      setWordDocumentLoading(false)
    } catch (error) {
      setWordDocumentError(`Error loading Word document: ${error.message}`)
      setWordDocumentLoading(false)
    }
  }

  const viewDocument = useCallback((doc) => {
    if (!doc?.url) {
      return message.error('Document URL is not available')
    }
    const url = doc.url
    setViewDocumentUrl(url)
    setExcelRendererData(null)
    setExcelRendererError(null)
    setWordDocumentContent(null)
    setWordDocumentError(null)

    const urlNoQuery = url.split('#')[0].split('?')[0]
    const ext = (urlNoQuery.split('.').pop() || '').toLowerCase()

    if (ext === 'xlsx' || ext === 'xls') {
      loadExcelWithRenderer(url)
    } else if (ext === 'docx' || ext === 'doc') {
      loadWordDocument(url)
    }
  }, [])

  // Refresh when master proposals are updated (optional sync)
  useEffect(() => {
    fetchPendingProposals()
  }, [fetchProposalsTrigger])

  const handleAcknowledge = async (id, acknowledge = true) => {
    setActionLoading((prev) => ({ ...prev, [id]: true }))
    try {
      const response = await fetch(`${API_BASE_URL}/proposals/acknowledge/${id}`, {
        method: 'PUT',
        headers: {
          accept: 'application/json',
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ is_acknowledged: acknowledge }),
      })

      if (!response.ok) {
        const err = await response.text()
        throw new Error(err || 'Failed to update acknowledgement')
      }

      message.success(acknowledge ? 'Proposal accepted' : 'Proposal rejected')
      fetchPendingProposals()  // Refresh pending list
      fetchProposalsTrigger()  // Refresh master proposals if needed
    } catch (error) {
      console.error(error)
      message.error(error.message || 'Failed to update acknowledgement')
    } finally {
      setActionLoading((prev) => ({ ...prev, [id]: false }))
    }
  }

  const pendingColumns = [

    {
      title: 'Enquiry Date', dataIndex: 'enquiry_date', key: 'enquiry_date', width: 120,
      render: (text) => formatDate(text)
    },

    { title: 'Customer Type', dataIndex: 'customer_type', key: 'customer_type', width: 120, },
    { title: 'Customer Name', dataIndex: 'customer_name', key: 'customer_name', width: 120, },
    { title: 'Address', dataIndex: 'address', key: 'address', width: 120, },
    { title: 'Email', dataIndex: 'email', key: 'email', width: 120, },
    { title: 'Phone No', dataIndex: 'phone_no', key: 'phone_no', width: 120, },
    { title: 'Alternate Contact Details', dataIndex: 'alternate_contact_details', key: 'alternate_contact_details', width: 120, },

    { title: 'Request Type', dataIndex: 'request_type', key: 'request_type', width: 120, },
    { title: 'Email Reference', dataIndex: 'email_reference', key: 'email_reference', width: 120, },
    { title: 'Quote Reference', dataIndex: 'quote_reference', key: 'quote_reference', width: 120, },

    { title: 'Quote Description', dataIndex: 'quote_description', key: 'quote_description', width: 120, },

    {
      title: 'Quote Date', dataIndex: 'quote_date', key: 'quote_date',
      render: (text) => formatDate(text), width: 120,
    },
    { title: 'Quote Amount', dataIndex: 'quote_amount', key: 'quote_amount', width: 120, },

    { title: 'Revised/Negotiated', dataIndex: 'revised/negotiated', key: 'revised/negotiated', width: 120, },
    {
      title: 'Revised/Negotiated Quote Date', dataIndex: 'revised/negotiated_quote_date', key: 'revised/negotiated_quote_date',
      render: (text) => formatDate(text), width: 120,
    },
    { title: 'Revised/Negotiated Quote Amount', dataIndex: 'revised/negotiated_quote_amount', key: 'revised/negotiated_quote_amount', width: 120, },

    { title: 'Quotation By', dataIndex: 'quotation_given_by_name', key: 'quotation_given_by_name', width: 120, },
    { title: 'Department', dataIndex: 'quotation_given_by_department', key: 'quotation_given_by_department', width: 120, },

    {
      key: 'enquiry_documents',
      title: 'Enquiry Documents',
      width: 160,
      render: (_, record) => {
        const count = record._docCount
        if (count === undefined) return <span style={{ color: '#999' }}>-</span>
        if (count > 0) {
          return (
            <Button type="link" size="small" icon={<EyeOutlined />} onClick={() => openDocsModal(record.id)}>
              View ({count})
            </Button>
          )
        }
        return <span style={{ color: '#999' }}>No documents</span>
      },
    },

   ...(!isGuest ? [ {
      title: 'Action',
      key: 'action',
      fixed: 'right',
      width: 150,
      render: (_, record) => (
        <Space size="small">
          <Button
            type="primary"
            size="small"
            loading={actionLoading[record.id]}
            onClick={() => handleAcknowledge(record.id, true)}
          >
            Accept
          </Button>
          <Popconfirm
            title="Reject this proposal?"
            description="This will acknowledge it as rejected."
            onConfirm={() => handleAcknowledge(record.id, false)}
            okText="Reject"
            okButtonProps={{ danger: true }}
          >
            <Button
              danger
              size="small"
              loading={actionLoading[record.id]}
            >
              Reject
            </Button>
          </Popconfirm>
        </Space>
      ),
    }] : []),
  ]

  return (
    <div className="overflow-x-auto">
      <Table
        rowKey="key"
        columns={pendingColumns}
        dataSource={pendingProposals}
        loading={loading}
        pagination={{ pageSize: 15 }}
        bordered
        scroll={{ x: 1600, y: 500 }}
        sticky
        locale={{ emptyText: 'No pending proposals to acknowledge' }}
      />

      {/* Uploaded Documents (Version List) Modal */}
      <Modal
        title="Uploaded Enquiry Documents"
        open={docsModalVisible}
        onCancel={() => setDocsModalVisible(false)}
        footer={[
          <Button key="close" onClick={() => setDocsModalVisible(false)}>Close</Button>,
        ]}
        width={700}
        maskClosable={false}
      >
        <Table
          rowKey="id"
          dataSource={projectDocs}
          loading={docsLoading}
          pagination={false}
          columns={[
            {
              title: 'Version',
              dataIndex: 'version',
              key: 'version',
              width: 80,
              render: (v) => (v ? `v${v}` : '-'),
            },
            {
              title: 'Name',
              dataIndex: 'display_name',
              key: 'name',
            },
            {
              title: 'Uploaded By',
              dataIndex: 'uploaded_by',
              key: 'uploaded_by',
              width: 150,
            },
            {
              title: 'Uploaded At',
              dataIndex: 'created_at',
              key: 'created_at',
              width: 180,
              render: (value) => (value ? dayjs(value).format('DD.MM.YYYY HH:mm') : '-'),
            },
            {
              title: 'View',
              key: 'view',
              width: 80,
              render: (_, record) => (
                <Button
                  type="link"
                  icon={<EyeOutlined />}
                  onClick={() => viewDocument(record)}
                />
              ),
            },
            {
              title: 'Attachments',
              dataIndex: 'attachment',
              key: 'attachment',
              width: 160,
              render: (attachments) => {
                if (!attachments || attachments.length === 0) {
                  return <span style={{ color: '#999' }}>-</span>
                }
                const VISIBLE_LIMIT = 2
                const visible = attachments.slice(0, VISIBLE_LIMIT)
                const overflow = attachments.slice(VISIBLE_LIMIT)

                const overflowContent = (
                  <Space direction="vertical" size={4}>
                    {overflow.map((url, idx) => (
                      <a
                        key={idx}
                        onClick={() => viewDocument({ url })}
                        style={{ fontSize: 12, cursor: 'pointer' }}
                      >
                        Attachment {VISIBLE_LIMIT + idx + 1}
                      </a>
                    ))}
                  </Space>
                )

                return (
                  <Space direction="vertical" size={2}>
                    {visible.map((url, idx) => (
                      <a
                        key={idx}
                        onClick={() => viewDocument({ url })}
                        style={{ fontSize: 12, cursor: 'pointer' }}
                      >
                        Attachment {idx + 1}
                      </a>
                    ))}
                    {overflow.length > 0 && (
                      <Popover
                        content={overflowContent}
                        title="More attachments"
                        trigger="click"
                        placement="right"
                      >
                        <a style={{ fontSize: 12, cursor: 'pointer' }}>
                          +{overflow.length} more
                        </a>
                      </Popover>
                    )}
                  </Space>
                )
              },
            },



          ]}
        />
        {(!docsLoading && !projectDocs.length) && (
          <div className="text-center text-gray-500 mt-4">No documents uploaded yet.</div>
        )}
      </Modal>

      <Modal
        title="Document Viewer"
        open={!!viewDocumentUrl}
        onCancel={() => {
          setViewDocumentUrl(null)
          setExcelRendererData(null)
          setExcelRendererError(null)
          setWordDocumentContent(null)
          setWordDocumentError(null)
        }}
        footer={null}
        width={1100}
      >
        {(() => {
          const currentUrl = viewDocumentUrl || ''
          const ext = (currentUrl.split('#')[0].split('?')[0].split('.').pop() || '').toLowerCase()

          if (!currentUrl) return null

          if (ext === 'xlsx' || ext === 'xls') {
            if (excelRendererLoading) return <div className="text-center py-10">Loading Excel file...</div>
            if (excelRendererError) {
              return (
                <div className="text-center py-10">
                  <p className="text-gray-500 mb-4">{excelRendererError}</p>
                  <Button type="primary" onClick={() => window.open(currentUrl, '_blank')}>Download Excel File</Button>
                </div>
              )
            }
            if (excelRendererData) {
              const rows = excelRendererData.rows || []
              const cols = excelRendererData.cols || []
              return rows.length > 0 ? (
                <div className="h-[70vh] overflow-auto border rounded">
                  <table style={{ borderCollapse: 'collapse', width: '100%', fontSize: 12 }}>
                    <thead>
                      <tr>
                        {cols.map((col, i) => (
                          <th key={i} style={{ border: '1px solid #d9d9d9', padding: 6, background: '#f5f5f5' }}>
                            {col.name || `Column ${i + 1}`}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {rows.map((row, ri) => (
                        <tr key={ri}>
                          {row.map((cell, ci) => (
                            <td key={ci} style={{ border: '1px solid #d9d9d9', padding: 6 }}>{cell}</td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : <div className="text-center py-10 text-gray-500">No data found in Excel file</div>
            }
            return <div className="text-center py-10">Loading Excel file...</div>
          }

          if (ext === 'docx' || ext === 'doc') {
            if (wordDocumentLoading) return <div className="text-center py-10">Loading Word document...</div>
            if (wordDocumentError) {
              return (
                <div className="text-center py-10">
                  <p className="text-gray-500 mb-4">{wordDocumentError}</p>
                  <Button type="primary" onClick={() => window.open(currentUrl, '_blank')}>Download Document</Button>
                </div>
              )
            }
            if (wordDocumentContent) {
              return (
                <div
                  className="h-[70vh] overflow-auto border rounded p-6 bg-white"
                  dangerouslySetInnerHTML={{ __html: wordDocumentContent }}
                />
              )
            }
            return <div className="text-center py-10">Loading Word document...</div>
          }

          return <iframe src={currentUrl} className="w-full h-[80vh]" title="Document" />
        })()}
      </Modal>
    </div>
  )
}


export default AcknowledgeProposalsTable;