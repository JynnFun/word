package com.jynn.mesh.controller;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import com.jynn.mesh.constant.AppConstant;
import com.jynn.mesh.util.WordUtil;

/**
 * 功能描述:word服务
 *
 * @author jynn
 * @created 2019年8月15日
 * @version 1.0.0
 */
@Controller
@RequestMapping("/word")
public class WordController {

	/**
	 * 功能描述:导出word
	 *
	 * @param request
	 * @param response
	 * @return
	 * @see [相关类/方法](可选)
	 * @since [产品/模块版本](可选)
	 */
	@RequestMapping("/export")
	@ResponseBody
	public void export(HttpServletRequest request, HttpServletResponse response) {
		
		//XWPFRun表示有相同属性的一段文本，所以模板里变量内容需要从左到右的顺序写，${name}，如果先写${},再添加内容，会拆分成几部分，不能正常使用

		// 病人数据
		Map<String, Object> patientMap = new HashMap<String, Object>();
		patientMap.put("${name}", "张三");
		patientMap.put("${relaId}", "ZZJ1234567");
		patientMap.put("${birthday}", "1996-12-31");
		patientMap.put("${sex}", "男");
		patientMap.put("${latestResult}", "心肌病(痰瘀互阻)");
		patientMap.put("${latestWestResult}", "心肌梗塞 Ⅱ期");

		// 病历数据
		List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
		// 项目标题
		List<List<String>> itemList = new ArrayList<List<String>>();

		Map<String, Object> map1 = new HashMap<String, Object>();
		map1.put("${title}", "第1次病历(2019-01-01)");
		map1.put("${result}", "心肌病");
		map1.put("${westResult}", "心肌梗塞 Ⅰ期");
		map1.put("${main}", "口干口苦");
		map1.put("${history}", "自感疲乏无力");
		map1.put("${pre}", "广藿香 10g  佩兰 10g  陈皮 10g");
		map1.put("${medInstruction}", "每日1剂");
		Map<String, Object> picture1 = new HashMap<String, Object>();
		picture1.put("width", 300);
		picture1.put("height", 400);
		picture1.put("type", "png");
		picture1.put("content",
				"https://wx4.sinaimg.cn/large/ae573a31ly1g63o7kt6kpj22c02c01hm.jpg;https://wx1.sinaimg.cn/large/0071e4B5gy1g63n3xff7rj32c0340b2a.jpg");
		map1.put("${image0}", picture1);
		Map<String, Object> picture2 = new HashMap<String, Object>();
		picture2.put("width", 300);
		picture2.put("height", 400);
		picture2.put("type", "png");
		picture2.put("content", "https://wx1.sinaimg.cn/large/ea5b1f20ly1g63k6fxpshj21kx1qy1kx.jpg");
		map1.put("${image1}", picture2);
		List<String> iList1 = new ArrayList<String>();
		iList1.add("血常规");
		iList1.add("尿常规");

		Map<String, Object> map2 = new HashMap<String, Object>();
		map2.put("${title}", "第2次病历(2019-08-18)");
		map2.put("${result}", "心肌病(痰瘀互阻)");
		map2.put("${westResult}", "心肌梗塞 Ⅱ期");
		map2.put("${main}", "口干口苦，口中粘腻减而未已");
		map2.put("${history}", "自感疲乏无力，口粘，甜腻");
		map2.put("${pre}", "广藿香 10g  佩兰 10g  陈皮 10g  法半夏 10g  太子参 10g");
		map2.put("${medInstruction}", "每日1剂 口服");
		Map<String, Object> picture3 = new HashMap<String, Object>();
		picture3.put("width", 300);
		picture3.put("height", 400);
		picture3.put("type", "png");
		picture3.put("content", "https://wx3.sinaimg.cn/large/005BYtFtly1g63epm0ad1j30jg0jgtb5.jpg");
		map2.put("${image0}", picture3);
		List<String> iList2 = new ArrayList<String>();
		iList2.add("血常规");

		list.add(map1);
		list.add(map2);

		itemList.add(iList1);
		itemList.add(iList2);

		WordUtil.DownloadWord(response, patientMap, list, itemList, AppConstant.TEMPLATE_PATH);
	}
}
