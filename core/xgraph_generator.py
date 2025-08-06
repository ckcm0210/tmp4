# graph_generator.py

import os
import webbrowser
from pyvis.network import Network
import json

class GraphGenerator:
    def __init__(self, nodes_data, edges_data):
        self.nodes_data = nodes_data
        self.edges_data = edges_data
        self.output_filename = "dependency_graph.html"

    def generate_graph(self):
        """
        使用 V5 方案（手動座標佈局 + 穩定互動）產生並打開 HTML 圖表。
        """
        # 1. 手動計算階層式佈局的初始座標
        self._calculate_node_positions()

        # 2. 使用 Pyvis 產生圖表
        net = Network(height="95vh", width="100%", bgcolor="#ffffff", font_color="black", directed=True)

        # 關鍵：不使用階層式佈局，直接設定全域選項，並啟用 HTML 支援
        options_str = """
        {
          "interaction": {
            "dragNodes": true,
            "dragView": true,
            "zoomView": true
          },
          "physics": {
            "enabled": false
          },
          "nodes": {
            "font": {
              "align": "left",
              "multi": "html"
            }
          },
          "edges": {
            "smooth": {
              "type": "cubicBezier",
              "forceDirection": "vertical",
              "roundness": 0.4
            }
          }
        }
        """
        net.set_options(options_str)

        # 將節點資料（包含計算好的 x, y 座標）加入網路圖
        for node_info in self.nodes_data:
            net.add_node(
                node_info["id"],
                label=node_info["label"],  # <-- FIX: Use the new 'label' key for default display
                shape='box',
                color=node_info["color"],
                x=node_info.get('x', 0),
                y=node_info.get('y', 0),
                fixed=False,
                title=node_info["title"],
                filename=node_info.get('filename', 'Current File'), # <<< --- THE FIX IS HERE
                # Pass all label variants to the node object for JS to use
                short_address_label=node_info["short_address_label"],
                full_address_label=node_info["full_address_label"],
                short_formula_label=node_info["short_formula_label"],
                full_formula_label=node_info["full_formula_label"],
                value_label=node_info["value_label"]
            )

        for edge in self.edges_data:
            net.add_edge(edge[0], edge[1])

        # 3. 注入 HTML 和 JavaScript
        temp_file = f"temp_{self.output_filename}"
        net.save_graph(temp_file)

        with open(temp_file, 'r', encoding='utf-8') as f:
            html_content = f.read()

        # --- 新的 HTML 注入 ---
        controls_html = """
        <div style='position: absolute; top: 10px; left: 10px; background: rgba(248, 249, 250, 0.95); padding: 12px; border: 1px solid #dee2e6; border-radius: 8px; z-index: 1000; font-family: sans-serif; font-size: 14px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);'>
          <div style='font-weight: bold; margin-bottom: 10px; color: #333;'>Display Options</div>
          
          <div style='margin-bottom: 8px;'>
            <label for='formulaToggle' style='cursor: pointer;'>
              <input type='checkbox' id='formulaToggle'> Show Full Formula Path
            </label>
          </div>
          
          <div style='margin-bottom: 8px;'>
            <label for='addressToggle' style='cursor: pointer;'>
              <input type='checkbox' id='addressToggle'> Show Full Address Path
            </label>
          </div>
          
          <div style='margin-bottom: 8px;'>
            <label for='fontSizeSlider' style='cursor: pointer; display: block; margin-bottom: 4px;'>
              Font Size: <span id='fontSizeValue'>14</span>px
            </label>
            <input type='range' id='fontSizeSlider' min='10' max='24' value='14' style='width: 100%;'>
          </div>
          
          <div style='border-top: 1px solid #ccc; margin-top: 12px; padding-top: 10px;'>
            <div style='font-weight: bold; margin-bottom: 6px; color: #333;'>File Legend</div>
            <div id='fileLegend' style='font-size: 12px;'>
              <!-- 動態生成的檔案圖例將在這裡顯示 -->
            </div>
            <div style='margin-top: 8px; font-size: 11px; color: #666;'>
              相同顏色 = 同一檔案<br>
              不同顏色間的箭頭 = 跨檔案依賴
            </div>
          </div>
        </div>
        """

        # --- 新的 JavaScript 注入 ---
        javascript_injection = """
        <script type='text/javascript'>
          document.addEventListener('DOMContentLoaded', function() {
            var network = window.network;
            var nodes = window.nodes;
            if (!network || !nodes) { return; }

            const formulaToggle = document.getElementById('formulaToggle');
            const addressToggle = document.getElementById('addressToggle');
            const fontSizeSlider = document.getElementById('fontSizeSlider');
            const fontSizeValue = document.getElementById('fontSizeValue');


            function updateNodeLabels() {
              const showFullAddress = addressToggle.checked;
              const showFullFormula = formulaToggle.checked;
              const fontSize = parseInt(fontSizeSlider.value);
              const currentPositions = network.getPositions();
              let updatedNodes = [];
              let allNodes = nodes.get({ returnType: 'Array' });

              allNodes.forEach(node => {
                const addressLabel = showFullAddress ? node.full_address_label : node.short_address_label;
                const formulaLabel = showFullFormula ? node.full_formula_label : node.short_formula_label;
                
                let newLabel = 'Address : <b>' + addressLabel + '</b>';
                
                if (formulaLabel && formulaLabel !== 'N/A') {
                  const displayFormula = formulaLabel.startsWith('=') ? formulaLabel : '=' + formulaLabel;
                  newLabel += '\\n\\nFormula : <i>' + displayFormula + '</i>';
                }
                
                newLabel += '\\n\\nValue     : ' + node.value_label;
                
                const position = currentPositions[node.id];
                if (position) {
                    const baseSize = 150;
                    const sizeMultiplier = fontSize / 14;
                    const nodeSize = Math.max(baseSize * sizeMultiplier, 100);
                    
                    
                    updatedNodes.push({
                      id: node.id,
                      label: newLabel,
                      x: position.x,
                      y: position.y,
                      fixed: true,
                      font: { size: fontSize, align: 'left' },
                      widthConstraint: { minimum: nodeSize, maximum: nodeSize * 1.5 },
                      heightConstraint: { minimum: nodeSize * 0.6, maximum: nodeSize * 1.5 }
                    });
                }
              });
              
              if (updatedNodes.length > 0) {
                  nodes.update(updatedNodes);
                  setTimeout(() => {
                      let releaseNodes = updatedNodes.map(n => ({ id: n.id, fixed: false }));
                      nodes.update(releaseNodes);
                  }, 100);
              }
            }

            function updateFontSize() {
              const fontSize = parseInt(fontSizeSlider.value);
              fontSizeValue.textContent = fontSize;
              updateNodeLabels();
            }

            function generateFileLegend() {
              const fileLegendDiv = document.getElementById('fileLegend');
              if (!fileLegendDiv) return;
              
              const fileColors = new Map();
              const allNodes = nodes.get({ returnType: 'Array' });
              
              allNodes.forEach(node => {
                // 從節點數據中讀取 filename 和 color
                const color = node.color || '#808080'; // 灰色作為後備
                const filename = node.filename || 'Unknown File'; // 後備名稱
                
                // 如果 Map 中沒有這個檔案名，就添加它
                if (!fileColors.has(filename)) {
                  fileColors.set(filename, color);
                }
              });
              
              // 對檔案名進行排序，確保 Current File 在最前面
              const sortedFiles = Array.from(fileColors.entries()).sort((a, b) => {
                if (a[0] === 'Current File') return -1;
                if (b[0] === 'Current File') return 1;
                return a[0].localeCompare(b[0]);
              });
              
              // 生成圖例的 HTML
              let legendHTML = '';
              sortedFiles.forEach(([filename, color]) => {
                legendHTML += `<div style="display: flex; align-items: center; margin-bottom: 4px;" title="檔案: ${filename}">`;
                legendHTML += `<div style="width: 16px; height: 16px; background-color: ${color}; margin-right: 8px; border-radius: 3px; border: 1px solid #ddd;"></div>`;
                legendHTML += `<span style="font-size: 12px; font-weight: 500; color: #333;">${filename}</span>`;
                legendHTML += '</div>';
              });
              
              // 將生成的 HTML 插入到圖例容器中
              fileLegendDiv.innerHTML = legendHTML;
            }

            addressToggle.addEventListener('change', updateNodeLabels);
            formulaToggle.addEventListener('change', updateNodeLabels);
            fontSizeSlider.addEventListener('input', updateFontSize);
            
            // 頁面載入後，立即生成檔案圖例
            generateFileLegend();
          });
        </script>
        """

        html_content = html_content.replace('<body>', '<body>\n' + controls_html)
        html_content = html_content.replace('</body>', javascript_injection + '\n</body>')


        final_file_path = os.path.join(os.getcwd(), self.output_filename)
        with open(final_file_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

        os.remove(temp_file)
        
        # 4. 在瀏覽器中打開
        webbrowser.open(f"file://{final_file_path}")
        print(f"Successfully generated interactive graph at: {final_file_path}")

    def _calculate_node_positions(self):
        """
        根據節點的層級（level），計算它們在圖表中的初始 x, y 座標。
        """
        level_counts = {}
        for node in self.nodes_data:
            level = node.get('level', 0)
            if level not in level_counts:
                level_counts[level] = 0
            level_counts[level] += 1

        level_y_step = 250
        level_x_step = 400  # 恢復正常間距

        current_level_counts = {level: 0 for level in level_counts}

        for node in self.nodes_data:
            level = node.get('level', 0)
            total_in_level = level_counts.get(level, 1)
            current_index_in_level = current_level_counts.get(level, 0)
            
            # 計算座標
            y = level * level_y_step
            x = (current_index_in_level - (total_in_level - 1) / 2.0) * level_x_step
            
            node['x'] = x
            node['y'] = y
            current_level_counts[level] = current_level_counts.get(level, 0) + 1
