import React, { useEffect, useMemo, useState } from "react";
import {
    Modal,
    Button,
    Radio,
    Input,
    InputNumber,
    Tag,
    Table,
    Tabs,
    Popconfirm,
    message,
    Empty,
} from "antd";
import { PlusOutlined, DeleteOutlined, FileWordOutlined } from "@ant-design/icons";
import axios from "axios";
import { API_BASE_URL } from '../config/api.js';

/* ============================================================
   CONSTANTS
   ============================================================ */

const MANPOWER_HEADER = "Manpower";
const MANPOWER_COLUMNS = ["Role", "Cost Breakup", "Amount"];
const DEFAULT_CUSTOM_COLUMNS = ["Description", "Amount"];

/* ============================================================
   API CALL - single stateless endpoint, nothing else needed.
   Swap this axios instance for your project's shared one if you have it.
   ============================================================ */

const api = axios.create({ baseURL: `${API_BASE_URL}/dynamic-tables` });

async function generateWordDocument(payload) {
    const res = await api.post("/generate-word", payload, { responseType: "blob" });

    const blobUrl = window.URL.createObjectURL(new Blob([res.data]));
    const disposition = res.headers["content-disposition"];
    let filename = "cost_breakdown.docx";
    if (disposition) {
        const match = disposition.match(/filename="?([^"]+)"?/);
        if (match) filename = match[1];
    }

    const link = document.createElement("a");
    link.href = blobUrl;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.URL.revokeObjectURL(blobUrl);
}

/* ============================================================
   HELPERS
   ============================================================ */

const emptyRow = (columns, headerName) => {
    const row = {};
    columns.forEach((col) => {
        if (headerName === MANPOWER_HEADER && col === "Cost Breakup") {
            row[col] = { rate: 0, hours: 0, days: 0, quantity: 1 };
        } else if (col === "Amount") {
            row[col] = 0;
        } else {
            row[col] = "";
        }
    });
    return row;
};

/* ============================================================
   SUB-COMPONENT: Manpower's 4-field inline cost input
   ============================================================ */

// function ManpowerCostInput({ value, onChange }) {
//   const rate = value?.rate ?? 0;
//   const hours = value?.hours ?? 0;
//   const days = value?.days ?? 0;
//   const quantity = value?.quantity ?? 1;

//   const update = (field, num) => {
//     onChange({ rate, hours, days, quantity, [field]: num ?? 0 });
//   };

//   return (
//     <div style={{ display: "flex", alignItems: "center", gap: 4, flexWrap: "wrap" }}>
//       <InputNumber min={0} value={rate} onChange={(v) => update("rate", v)} placeholder="Rate" style={{ width: 80 }} />
//       <span>*</span>
//       <InputNumber min={0} value={hours} onChange={(v) => update("hours", v)} placeholder="Hrs" style={{ width: 70 }} />
//       <span style={{ color: "#888" }}>(hours) *</span>
//       <InputNumber min={0} value={days} onChange={(v) => update("days", v)} placeholder="Days" style={{ width: 70 }} />
//       <span style={{ color: "#888" }}>(days) *</span>
//       <InputNumber min={0} value={quantity} onChange={(v) => update("quantity", v)} placeholder="Qty" style={{ width: 60 }} />
//     </div>
//   );
// }


function ManpowerCostInput({ value, onChange }) {
    const rate = value?.rate ?? 0;
    const hours = value?.hours ?? 0;
    const days = value?.days ?? 0;
    const quantity = value?.quantity ?? 0;
    const amount = rate * hours * days * quantity;

    const update = (field, num) => {
        onChange({ rate, hours, days, quantity, [field]: num ?? 0 });
    };

    const field = (label, key, val, width) => (
        <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
            <span style={{ fontSize: 11, color: "#888", fontWeight: 500 }}>{label}</span>
            <InputNumber
                min={0}
                value={val === 0 ? undefined : val}
                onChange={(v) => update(key, v)}
                placeholder="0"
                style={{ width }}
            />
        </div>
    );

    return (
        <div style={{ display: "flex", alignItems: "flex-end", gap: 8, flexWrap: "wrap" }}>
            {field("Rate (₹/unit)", "rate", rate, 90)}
            <span style={{ marginBottom: 6, color: "#bbb" }}>×</span>
            {field("Hours", "hours", hours, 70)}
            <span style={{ marginBottom: 6, color: "#bbb" }}>×</span>
            {field("Days", "days", days, 70)}
            <span style={{ marginBottom: 6, color: "#bbb" }}>×</span>
            {field("Quantity", "quantity", quantity, 70)}
        </div>
    );
}

/* ============================================================
   SUB-COMPONENT: editable rows for one header (in-memory only)
   ============================================================ */

function HeaderRowsEditor({ headerItem, onChange, onNewTable }) {
    const { header_name: headerName, columns, rows } = headerItem;
    const [addingColumn, setAddingColumn] = useState(false);
    const [newColumnName, setNewColumnName] = useState("");

    const updateRow = (index, key, value) => {
        const next = [...rows];
        next[index] = { ...next[index], [key]: value };
        onChange({ ...headerItem, rows: next });
    };

    const addRow = () => onChange({ ...headerItem, rows: [...rows, emptyRow(columns, headerName)] });
    const removeRow = (index) => onChange({ ...headerItem, rows: rows.filter((_, i) => i !== index) });

    const confirmAddColumn = () => {
        const trimmed = newColumnName.trim();
        if (!trimmed) {
            setAddingColumn(false);
            return;
        }
        if (columns.includes(trimmed)) {
            message.warning("Column already exists");
            return;
        }

        // Insert right before "Amount" if it exists, otherwise append at the end.
        const amountIndex = columns.indexOf("Amount");
        const nextColumns =
            amountIndex === -1
                ? [...columns, trimmed]
                : [...columns.slice(0, amountIndex), trimmed, ...columns.slice(amountIndex)];

        const nextRows = rows.map((r) => ({ ...r, [trimmed]: "" }));
        onChange({ ...headerItem, columns: nextColumns, rows: nextRows });
        setNewColumnName("");
        setAddingColumn(false);
    };

    const previewTotal = useMemo(() => {
        if (headerName === MANPOWER_HEADER) {
            return rows.reduce((sum, r) => {
                const cb = r["Cost Breakup"] || {};
                return sum + (cb.rate || 0) * (cb.hours || 0) * (cb.days || 0) * (cb.quantity || 1);
            }, 0);
        }
        if (columns.includes("Amount")) {
            return rows.reduce((sum, r) => sum + (Number(r["Amount"]) || 0), 0);
        }
        return null;
    }, [rows, columns, headerName]);
    const tableColumns = [
        ...columns.map((col) => ({
            title: col,
            dataIndex: col,
            key: col,
            render: (_, record, index) => {
                if (headerName === MANPOWER_HEADER && col === "Cost Breakup") {
                    return <ManpowerCostInput value={record[col]} onChange={(v) => updateRow(index, col, v)} />;
                }
                if (headerName === MANPOWER_HEADER && col === "Amount") {
                    const cb = record["Cost Breakup"] || {};
                    const amt = (cb.rate || 0) * (cb.hours || 0) * (cb.days || 0) * (cb.quantity || 0);
                    return <span>{amt.toFixed(2)}</span>;
                }
                if (col === "Amount") {
                    return (
                        <InputNumber min={0} value={record[col]} onChange={(v) => updateRow(index, col, v ?? 0)} style={{ width: 120 }} />
                    );
                }
                return <Input value={record[col]} onChange={(e) => updateRow(index, col, e.target.value)} />;
            },
        })),


        ...(headerName === MANPOWER_HEADER
            ? []
            : [
                {
                    title: addingColumn ? (
                        <div style={{ display: "flex", gap: 4 }} onClick={(e) => e.stopPropagation()}>
                            <Input
                                size="small"
                                autoFocus
                                placeholder="Column name"
                                value={newColumnName}
                                onChange={(e) => setNewColumnName(e.target.value)}
                                onPressEnter={confirmAddColumn}
                                style={{ width: 110 }}
                            />
                            <Button size="small" type="primary" onClick={confirmAddColumn}>
                                Add
                            </Button>
                        </div>
                    ) : (
                        <Button
                            size="small"
                            type="text"
                            icon={<PlusOutlined />}
                            onClick={() => setAddingColumn(true)}
                            title="Add column"
                        />
                    ),
                    key: "__add_column__",
                    width: addingColumn ? 190 : 50,
                },
            ]),
        {
            title: "",
            key: "actions",
            width: 50,
            render: (_, __, index) => (
                <Popconfirm title="Remove row?" onConfirm={() => removeRow(index)}>
                    <Button type="text" danger icon={<DeleteOutlined />} />
                </Popconfirm>
            ),
        },
    ];

    return (
        <Table
            rowKey={(_, index) => String(index)}
            columns={tableColumns}
            dataSource={rows}
            pagination={false}
            bordered
            size="small"
            footer={() => (
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <div style={{ display: "flex", gap: 8 }}>
                        <Button icon={<PlusOutlined />} onClick={addRow}>
                            Add Row
                        </Button>
                        <Button icon={<PlusOutlined />} onClick={onNewTable}>
                            New Table
                        </Button>
                    </div>
                    {previewTotal !== null && <span style={{ fontWeight: 600 }}>Total: {previewTotal.toFixed(2)}</span>}
                </div>
            )}
        />
    );
}

function AddHeaderForm({ existingHeaderNames, onAdd, isEnteringHeader, activeHeaderName }) {
    const [customName, setCustomName] = useState("");

    const handleAddCustom = () => {
        const name = customName.trim();
        if (!name) {
            message.warning("Enter a header name");
            return;
        }
        if (existingHeaderNames.includes(name)) {
            message.warning("A header with this name already exists");
            return;
        }
        onAdd({
            header_name: name,
            columns: DEFAULT_CUSTOM_COLUMNS,
            rows: [emptyRow(DEFAULT_CUSTOM_COLUMNS, name)],
        });
        setCustomName("");
    };

    if (isEnteringHeader) {
        return (
            <div style={{ border: "1px dashed #ccc", padding: 12, borderRadius: 6, marginBottom: 16 }}>
                <div style={{ display: "flex", gap: 8 }}>
                    <Input
                        autoFocus
                        placeholder="Enter Table Name"
                        value={customName}
                        onChange={(e) => setCustomName(e.target.value)}
                        onPressEnter={handleAddCustom}
                    />
                    <Button type="primary" icon={<PlusOutlined />} onClick={handleAddCustom}>
                        Create
                    </Button>
                </div>
            </div>
        );
    }

    // Only show the Manpower explanation while the Manpower tab is active -
    // stay silent for any other (custom) table's tab.
    if (activeHeaderName !== MANPOWER_HEADER) {
        return null;
    }

    return (
        <div style={{ fontSize: 13, color: "#666", marginBottom: 16 }}>
            Role, Cost Breakup (rate × hours × days × quantity), Amount - calculated automatically.
        </div>
    );
}

export function CostEstimationModal({ open, onClose, title, createdBy }) {
    const [headers, setHeaders] = useState([]); // [{header_name, columns, rows}]
    const [activeKey, setActiveKey] = useState("");
    const [generating, setGenerating] = useState(false);
    const [isEnteringHeader, setIsEnteringHeader] = useState(false);

    // Reset form state each time the modal is freshly opened, and
    // auto-create the Manpower table right away - no click required.
    useEffect(() => {
        if (open) {
            const initial = {
                header_name: MANPOWER_HEADER,
                columns: MANPOWER_COLUMNS,
                rows: [emptyRow(MANPOWER_COLUMNS, MANPOWER_HEADER)],
            };
            setHeaders([initial]);
            setActiveKey(MANPOWER_HEADER);
            setIsEnteringHeader(false);
        } else {
            setHeaders([]);
            setActiveKey("");
            setIsEnteringHeader(false);
        }
    }, [open]);

    // Switching TO custom mode means the user is about to type a header name -
    // hide existing tables until it's actually added.
    const handleFormModeChange = (newMode) => {
        setFormMode(newMode);
        setIsEnteringHeader(newMode === "custom");
    };

    const closeModal = () => onClose();

    const handleAddHeader = (item) => {
        setHeaders((prev) => [...prev, item]);
        setActiveKey(item.header_name);
        setIsEnteringHeader(false);
    };

    const handleHeaderChange = (index, updated) => {
        setHeaders((prev) => prev.map((h, i) => (i === index ? updated : h)));
    };

    const handleRemoveHeader = (headerName) => {
        setHeaders((prev) => prev.filter((h) => h.header_name !== headerName));
    };

    const handleGenerate = async () => {
        if (headers.length === 0) {
            message.warning("Add at least one header before generating");
            return;
        }
        setGenerating(true);
        try {
            await generateWordDocument({
                title: title || "Cost Breakdown",
                created_by: createdBy,
                tables: headers,
            });
            message.success("Word document generated");
            closeModal();
        } catch (err) {
            message.error("Failed to generate document");
        } finally {
            setGenerating(false);
        }
    };

    const existingHeaderNames = headers.map((h) => h.header_name);

    return (
        <Modal
            title={title ? `Cost Estimation - ${title}` : "Cost Estimation"}
            open={open}
            onCancel={closeModal}
            width={800}
            destroyOnClose
            footer={[
                <Button key="cancel" onClick={closeModal}>
                    Cancel
                </Button>,
                <Button
                    key="generate"
                    type="primary"
                    icon={<FileWordOutlined />}
                    loading={generating}
                    onClick={handleGenerate}
                >
                    Generate Word Document
                </Button>,
            ]}
        >
            <AddHeaderForm
                existingHeaderNames={existingHeaderNames}
                onAdd={handleAddHeader}
                isEnteringHeader={isEnteringHeader}
                activeHeaderName={activeKey}
            />

            {isEnteringHeader ? null : headers.length === 0 ? (
                <Empty description="No headers added yet" />
            ) : (
                <Tabs
                    activeKey={activeKey}
                    onChange={setActiveKey}
                    type="editable-card"
                    hideAdd
                    onEdit={(targetKey, action) => {
                        if (action === "remove") handleRemoveHeader(targetKey);
                    }}
                    items={headers.map((h, index) => ({
                        key: h.header_name,
                        label: h.header_name,
                        children: (
                            <HeaderRowsEditor
                                headerItem={h}
                                onChange={(updated) => handleHeaderChange(index, updated)}
                                onNewTable={() => setIsEnteringHeader(true)}
                            />
                        ),
                    }))}
                />
            )}
        </Modal>
    );
}

//export { CostEstimationModal };