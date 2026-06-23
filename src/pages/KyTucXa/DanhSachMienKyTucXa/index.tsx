import TableBase from '@/components/Table';
import ButtonExtend from '@/components/Table/ButtonExtend';
import type { IColumn } from '@/components/Table/typing';
import FilterHocKy from '@/pages/DaoTaoV2/HocKy/HocKy/components/FilterHocKy';
import { DeleteOutlined, EditOutlined } from '@ant-design/icons';
import { Popconfirm, message } from 'antd';
import fileDownload from 'js-file-download';
import { useState } from 'react';
import { useIntl, useModel } from 'umi';
import * as XLSX from 'xlsx';
import FormSinhVien from './components/FormSinhVien';

const DanhSachMienKyTucXa = () => {
	const intl = useIntl();
	const t = (id: string, values?: Record<string, any>) => intl.formatMessage({ id }, values);
	const { page, limit, handleEdit, deleteModel, getModel, postManyByMaHocKy, danhSach, currentDanhSachMien } =
		useModel('kytucxa.danhsachmiensinhvien');
	const { record: recHocKy } = useModel('daotaov2.hocky.hocky');
	const [visibleSelect, setVisibleSelect] = useState(false);

	const getData = () => {
		if (recHocKy?.ma) getModel({ maHocKy: recHocKy?.ma });
	};

	const renderText = (value?: string | number | null) => value || '-';

	const getKhoaNganh = (record: any) => {
		const khoaNganh = record?.khoaNganh;
		const sinhVienKhoaNganh = record?.sinhVien?.khoaNganh;

		return (
			khoaNganh?.ten ||
			khoaNganh?.ma ||
			(typeof khoaNganh === 'string' ? khoaNganh : undefined) ||
			record?.tenKhoaNganh ||
			record?.maKhoaNganh ||
			sinhVienKhoaNganh?.ten ||
			sinhVienKhoaNganh?.ma ||
			(typeof sinhVienKhoaNganh === 'string' ? sinhVienKhoaNganh : undefined) ||
			record?.sinhVien?.tenKhoaNganh ||
			record?.sinhVien?.maKhoaNganh
		);
	};

	const getSoDienThoai = (record: any) =>
		record?.soDienThoai || record?.sdt || record?.sinhVien?.soDienThoai || record?.sinhVien?.sdt;

	const getEmail = (record: any) => record?.email || record?.sinhVien?.email;

	const handleAddStudentsDone = async (
		newStudents: {
			maSinhVien: string;
			hoTen: string;
			khoaSinhVien: string;
			khoaNganh?: string;
			soDienThoai?: string;
			email?: string;
		}[],
	) => {
		if (!recHocKy?.ma) {
			message.warning(t('kytucxa.danhsachmien.message.selectSemesterFirst'));
			return;
		}

		await postManyByMaHocKy(
			recHocKy.ma,
			newStudents.map((student) => ({
				maSinhVien: student.maSinhVien,
				hoTen: student.hoTen,
				khoaSinhVien: student.khoaSinhVien || '',
				khoaNganh: student.khoaNganh || '',
				soDienThoai: student.soDienThoai || '',
				email: student.email || '',
			})),
		);
		setVisibleSelect(false);
	};

	const handleExportExcel = () => {
		if (!danhSach?.length) {
			message.warning(t('kytucxa.danhsachmien.message.noDataToExport'));
			return;
		}

		const dataToExport = danhSach.map((item: any, index: number) => ({
			TT: index + 1,
			[t('kytucxa.danhsachmien.maSinhVien')]: item.code || '',
			[t('kytucxa.danhsachmien.hoTen')]: item.fullname || '',
			[t('kytucxa.danhsachmien.khoaSinhVien')]: item.khoaSinhVien || '',
			[t('kytucxa.danhsachmien.khoaNganh')]: getKhoaNganh(item) || '',
			[t('kytucxa.danhsachmien.soDienThoai')]: getSoDienThoai(item) || '',
			Email: getEmail(item) || '',
			[t('kytucxa.danhsachmien.trangThaiMinhChung')]: item.trangThaiMinhChung || t('kytucxa.danhsachmien.choDuyet'),
		}));

		const worksheet = XLSX.utils.json_to_sheet(dataToExport);
		const workbook = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(workbook, worksheet, t('kytucxa.danhsachmien.excel.sheetDanhSach'));
		const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
		const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
		fileDownload(blob, t('kytucxa.danhsachmien.excel.fileName', { maHocKy: recHocKy?.ma || '' }));
	};

	const columns: IColumn<any>[] = [
		{
			title: t('kytucxa.danhsachmien.maSV'),
			dataIndex: 'code',
			width: 120,
			filterType: 'string',
			render: (text: string) => <strong>{text}</strong>,
		},
		{
			title: t('kytucxa.danhsachmien.hoTen'),
			dataIndex: 'fullname',
			width: 180,
			filterType: 'string',
		},
		{
			title: t('kytucxa.danhsachmien.khoa'),
			dataIndex: 'khoaSinhVien',
			width: 120,
		},
		{
			title: t('kytucxa.danhsachmien.khoaNganh'),
			dataIndex: 'khoaNganh',
			width: 180,
			render: (_value, record) => renderText(getKhoaNganh(record)),
		},
		{
			title: t('kytucxa.danhsachmien.soDienThoai'),
			dataIndex: 'soDienThoai',
			width: 130,
			filterType: 'string',
			render: (_value, record) => renderText(getSoDienThoai(record)),
		},
		{
			title: 'Email',
			dataIndex: 'email',
			width: 200,
			filterType: 'string',
			render: (_value, record) => renderText(getEmail(record)),
		},
		{
			title: t('kytucxa.danhsachmien.minhChung'),
			dataIndex: 'urlMinhChung',
			key: 'urlMinhChung',
			width: 180,
			render: (val: string) => {
				if (!val) return <span style={{ color: '#bfbfbf' }}>{t('kytucxa.danhsachmien.chuaNop')}</span>;

				const filename = val.substring(val.lastIndexOf('/') + 1) || 'minh-chung.pdf';
				return (
					<a
						href={val}
						target='_blank'
						rel='noopener noreferrer'
						style={{ textDecoration: 'underline', color: '#125195', fontWeight: 500 }}
					>
						{filename}
					</a>
				);
			},
			fixed: 'right',
		},
		{
			title: t('kytucxa.danhsachmien.thaoTac'),
			key: 'action',
			width: 100,
			align: 'center',
			fixed: 'right',
			render: (_value, record) => (
				<>
					<ButtonExtend
						tooltip={t('global.button.chinhsua')}
						onClick={() => handleEdit(record)}
						type='link'
						icon={<EditOutlined />}
					/>
					<Popconfirm
						onConfirm={() => deleteModel(record._id, getData)}
						title={t('kytucxa.danhsachmien.confirmDeleteStudent')}
						placement='topRight'
					>
						<ButtonExtend tooltip={t('global.button.xoa')} danger type='link' icon={<DeleteOutlined />} />
					</Popconfirm>
				</>
			),
		},
	];

	return (
		<>
			<TableBase
				getData={getData}
				columns={columns}
				dependencies={[page, limit, recHocKy?.ma]}
				modelName='kytucxa.danhsachmiensinhvien'
				title={t('kytucxa.danhsachmien.title')}
				Form={FormSinhVien}
				formProps={{ danhSachId: currentDanhSachMien?._id, maHocKy: recHocKy?.ma, getData }}
				scroll={{ x: 1300 }}
				buttons={{
					import: true,
					export: true,
					create: !!currentDanhSachMien?._id,
				}}
				// otherButtons={[
				// 	<ButtonExtend
				// 		key='btn-import-student'
				// 		icon={<ImportOutlined />}
				// 		disabled={!currentDanhSachMien?._id}
				// 		onClick={() => setVisibleSelect(true)}
				// 	>
				// 		{t('global.button.nhapdulieu')}
				// 	</ButtonExtend>,
				// 	<ButtonExtend key='btn-export-student' icon={<ExportOutlined />} onClick={handleExportExcel}>
				// 		{t('global.button.xuatdulieu')}
				// 	</ButtonExtend>,
				// ]}
				showModalTitle
			>
				<div style={{ marginBottom: 12 }}>
					<FilterHocKy isSetHocKy width={300} hideExpand />
				</div>
			</TableBase>

			{/* <StudentSelectModal
				open={visibleSelect}
				onCancel={() => setVisibleSelect(false)}
				activeSemester={currentDanhSachMien as any}
				selectedSemesterMa={recHocKy?.ma}
				existingStudents={danhSach || []}
				onOk={handleAddStudentsDone}
			/> */}
		</>
	);
};

export default DanhSachMienKyTucXa;
