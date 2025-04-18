class ChordTransposer:
    # 所有可能的调性
    KEYS = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']
    
    # 和弦音符映射
    CHORD_NOTES = {
        'C': 0, 'C#': 1, 'Db': 1,
        'D': 2, 'D#': 3, 'Eb': 3,
        'E': 4,
        'F': 5, 'F#': 6, 'Gb': 6,
        'G': 7, 'G#': 8, 'Ab': 8,
        'A': 9, 'A#': 10, 'Bb': 10,
        'B': 11
    }
    
    @classmethod
    def normalize_note(cls, note: str) -> str:
        """规范化音符名称，处理多重升降号"""
        if not note:
            return note
            
        # 提取基本音符和升降号
        base_note = note[0]  # 第一个字符一定是基本音符
        accidentals = note[1:]  # 剩余部分是升降号
        
        print(f"[normalize_note] 输入音符: {note}")
        print(f"[normalize_note] 基本音符: {base_note}")
        print(f"[normalize_note] 升降号: {accidentals}")
        
        # 计算升降号的总效果
        semitones = 0
        for acc in accidentals:
            if acc == '#':
                semitones += 1
            elif acc == 'b':
                semitones -= 1
        
        print(f"[normalize_note] 升降半音数: {semitones}")
        
        # 计算最终音高
        base_num = cls.CHORD_NOTES.get(base_note, 0)
        final_num = (base_num + semitones) % 12
        
        result = cls.KEYS[final_num]
        print(f"[normalize_note] 基本音符数值: {base_num}")
        print(f"[normalize_note] 最终音符数值: {final_num}")
        print(f"[normalize_note] 规范化结果: {result}")
        
        return result
    
    @classmethod
    def get_key_number(cls, key: str) -> int:
        """获取调号的数字表示"""
        print(f"[get_key_number] 输入调号: {key}")
        # 首先规范化音符
        normalized_key = cls.normalize_note(key)
        result = cls.CHORD_NOTES.get(normalized_key, 0)
        print(f"[get_key_number] 规范化后: {normalized_key}")
        print(f"[get_key_number] 调号数值: {result}")
        return result
    
    @classmethod
    def transpose_chord(cls, chord: str, from_key: str, to_key: str) -> str:
        """转调单个和弦"""
        print(f"\n[transpose_chord] 开始转调和弦: {chord}")
        print(f"[transpose_chord] 从调号: {from_key} 到调号: {to_key}")
        
        if not chord or len(chord) < 1:
            print("[transpose_chord] 空和弦，直接返回")
            return chord
            
        # 计算调号差
        from_num = cls.get_key_number(from_key)
        to_num = cls.get_key_number(to_key)
        
        # 修改计算半音数的方式
        semitones = to_num - from_num
        if semitones < 0:
            semitones += 12
        
        print(f"[transpose_chord] 从 {from_num} 到 {to_num}")
        print(f"[transpose_chord] 需要升高的半音数: {semitones}")
        
        # 分离和弦的根音和修饰符
        root = ''
        modifier = ''
        bass = ''
        
        # 处理带有低音的和弦（例如 C/G）
        if '/' in chord:
            main_chord, bass = chord.split('/')
            print(f"[transpose_chord] 带低音和弦 - 主和弦: {main_chord}, 低音: {bass}")
        else:
            main_chord = chord
            print(f"[transpose_chord] 普通和弦: {main_chord}")
            
        # 提取和弦根音和修饰符
        i = 0
        while i < len(main_chord):
            if main_chord[i].isalpha():
                root += main_chord[i]
                i += 1
                # 收集所有的升降号
                while i < len(main_chord) and main_chord[i] in ['#', 'b']:
                    root += main_chord[i]
                    i += 1
                modifier = main_chord[i:] if i < len(main_chord) else ''
                break
            i += 1
                
        print(f"[transpose_chord] 解析结果 - 根音: {root}, 修饰符: {modifier}")
        
        # 如果无法识别和弦，返回原始字符串
        if not root:
            print("[transpose_chord] 无法识别和弦，返回原始字符串")
            return chord
            
        # 规范化并转调根音
        normalized_root = cls.normalize_note(root)
        root_num = cls.CHORD_NOTES[normalized_root]
        new_root_num = (root_num + semitones) % 12
        new_root = cls.KEYS[new_root_num]
        
        print(f"[transpose_chord] 根音转调 - 规范化: {normalized_root}, 新根音: {new_root}")
        
        # 转调低音（如果存在）
        if bass:
            normalized_bass = cls.normalize_note(bass)
            bass_num = cls.CHORD_NOTES[normalized_bass]
            new_bass_num = (bass_num + semitones) % 12
            new_bass = cls.KEYS[new_bass_num]
            result = f"{new_root}{modifier}/{new_bass}"
            print(f"[transpose_chord] 最终带低音和弦: {result}")
            return result
            
        result = f"{new_root}{modifier}"
        print(f"[transpose_chord] 最终和弦: {result}")
        return result
    
    @classmethod
    def transpose_text(cls, text: str, from_key: str, to_key: str, preserve_spaces: bool = True) -> str:
        """
        转调文本中的所有和弦
        
        Args:
            text: 要转调的文本
            from_key: 原始调号
            to_key: 目标调号
            preserve_spaces: 是否保留原始空格，默认为True
        """
        print(f"\n[transpose_text] 开始转调文本")
        print(f"[transpose_text] 原始文本: {text}")
        print(f"[transpose_text] 从调号: {from_key} 到调号: {to_key}")
        print(f"[transpose_text] 保留空格: {preserve_spaces}")
        
        if not text or not from_key or not to_key:
            print("[transpose_text] 空文本或缺少调号信息，直接返回")
            return text
            
        if preserve_spaces:
            # 保留所有空格，包括连续空格
            # 使用split(' ')而不是split()以保留连续空格
            words = text.split(' ')
            # 如果文本以空格开头或结尾，split会产生空字符串，我们需要保留这些
            new_words = []
            for i, word in enumerate(words):
                if word == '':
                    new_words.append('')
                elif word and word[0].isupper() and any(note in word for note in cls.CHORD_NOTES.keys()):
                    new_words.append(cls.transpose_chord(word, from_key, to_key))
                else:
                    new_words.append(word)
            # 使用' '.join以保留原始空格
            result = ' '.join(new_words)
            print(f"[transpose_text] 转调后文本: {result}")
            return result
        else:
            # 不保留连续空格的原始行为
            words = text.split()
            new_words = []
            for word in words:
                if word and word[0].isupper() and any(note in word for note in cls.CHORD_NOTES.keys()):
                    new_words.append(cls.transpose_chord(word, from_key, to_key))
                else:
                    new_words.append(word)
            result = ' '.join(new_words)
            print(f"[transpose_text] 转调后文本: {result}")
            return result 