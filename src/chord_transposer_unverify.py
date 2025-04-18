class ChordTransposer:
    def __init__(self):
        self.notes = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']
        self.flat_notes = ['C', 'Db', 'D', 'Eb', 'E', 'F', 'Gb', 'G', 'Ab', 'A', 'Bb', 'B']
        
    def get_key_number(self, key: str) -> int:
        """獲取音符的數字表示（0-11）"""
        key = key.strip()
        if not key:
            return 0
            
        # 處理降號
        if 'b' in key:
            for i, note in enumerate(self.flat_notes):
                if key == note:
                    return i
        # 處理升號
        else:
            for i, note in enumerate(self.notes):
                if key == note:
                    return i
        return 0

    def transpose_chord(self, chord: str, from_key: str, to_key: str) -> str:
        """轉調和弦"""
        if not chord or not from_key or not to_key:
            return chord

        # 計算調號差
        from_num = self.get_key_number(from_key)
        to_num = self.get_key_number(to_key)
        semitones = (to_num - from_num) % 12

        # 分離和弦的各個部分
        parts = []
        current = ''
        for char in chord:
            if char.isalpha() or char in '#b':
                current += char
            else:
                if current:
                    parts.append(current)
                    current = ''
                parts.append(char)
        if current:
            parts.append(current)

        # 轉調每個部分
        result = ''
        i = 0
        while i < len(parts):
            part = parts[i]
            # 如果是音符
            if part[0].isalpha():
                # 檢查是否是根音或低音音符
                is_note = False
                for note in self.notes + self.flat_notes:
                    if part.startswith(note):
                        is_note = True
                        break
                
                if is_note:
                    # 獲取原始音符的索引
                    note_index = -1
                    if 'b' in part:
                        for idx, note in enumerate(self.flat_notes):
                            if part.startswith(note):
                                note_index = idx
                                break
                    else:
                        for idx, note in enumerate(self.notes):
                            if part.startswith(note):
                                note_index = idx
                                break

                    if note_index != -1:
                        # 計算新音符
                        new_index = (note_index + semitones) % 12
                        # 使用升號記號
                        new_note = self.notes[new_index]
                        # 添加剩餘的修飾符（如 m7）
                        result += new_note + part[len(new_note):]
                    else:
                        result += part
                else:
                    result += part
            else:
                result += part
            i += 1

        return result

    def transpose_text(self, text: str, from_key: str, to_key: str) -> str:
        """轉調文本中的和弦"""
        if not text or not from_key or not to_key:
            return text

        # 分割文本，保留空格
        parts = text.split(' ')
        result = []
        
        for part in parts:
            # 檢查是否是和弦（包含音符和可能的修飾符）
            is_chord = False
            for note in self.notes + self.flat_notes:
                if part.startswith(note):
                    is_chord = True
                    break
            
            if is_chord:
                # 處理斜槓和弦（如 B/D#）
                if '/' in part:
                    bass_parts = part.split('/')
                    if len(bass_parts) == 2:
                        root = self.transpose_chord(bass_parts[0], from_key, to_key)
                        bass = self.transpose_chord(bass_parts[1], from_key, to_key)
                        result.append(f"{root}/{bass}")
                    else:
                        result.append(part)
                else:
                    result.append(self.transpose_chord(part, from_key, to_key))
            else:
                result.append(part)
        
        return ' '.join(result) 