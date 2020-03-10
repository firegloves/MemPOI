/**
 * created by firegloves
 */

package it.firegloves.mempoi.domain.pivottable;

import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.MempoiTable;
import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.util.AreaReference;

@Data
@Accessors(chain = true)
public class MempoiPivotTableSource {

    // TODO capire se è meglio usare direttamente la table => non credo perché in fase di configurazione ancora non è stata creata
    private final MempoiTable mempoiTable;
    private final AreaReference areaReference;
    private final MempoiSheet mempoiSheet;
//      mempoisheet sorgente. Se null => mempoi sheet in cui viene inserita la pivottable
}