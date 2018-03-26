package edu.zjnu.excelstatistics.utils;

import java.util.Vector;

import javax.swing.AbstractListModel;

public class DataModel extends AbstractListModel {
		private Vector data = null;

		public DataModel(Object[] listData) {
			data = new Vector();
			for (int i = 0; i < listData.length; i++) {
				data.add(listData[i]);
			}
		}

		public void addData(Object o) {
			data.add(o);
		}

		public void removeIndex(int index) {
			data.remove(index);
		}

		public int getSize() {
			return data.size();
		}

		public Object getElementAt(int index) {
			return data.get(index);
		}
	}